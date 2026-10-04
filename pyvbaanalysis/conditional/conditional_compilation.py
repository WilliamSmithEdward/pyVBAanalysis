"""Conditional-compilation activity tracking (MS-VBAL 3.4).

Ported from xlide_vscode/src/analyzer/conditional/conditionalCompilation.ts.
Replays the #If/#ElseIf/#Else/#End If directive stack with the default compiler
constants of 64-bit Office on Windows (VBA7, VBA6, Win64 and Win32 are 1; Win16
and Mac are 0) plus any project #Const definitions, and reports whether a given
source span is active, inactive, or unknown. Unknown stays unknown: the analyzer
never guesses a branch.
"""

from __future__ import annotations

import enum
import math
import re
import unicodedata
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Protocol, Union

from ..constants.date_literal import date_literal_serial
from ..constants.integer_constant_expression import bankers_round, parse_vba_integer_literal
from ..js_compat import js_number, js_trim
from ..lexer.token_helpers import relational_operator_at, token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from ..parser.nodes import (
    BodyNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    EnumNode,
    ModuleNode,
    ProcedureNode,
    Span,
    TypeNode,
    iter_body_nodes,
)

@dataclass(frozen=True, slots=True)
class ConditionalSpecialValue:
    """The values a directive expression has beside the plain ones: Empty, Null,
    Nothing and a Date, which is its serial (XLIDE issue #208). ``kind`` is
    "empty", "null", "nothing" or "date"; ``serial`` is set for a Date only."""

    kind: str
    serial: float | None = None


# A resolved conditional value. bool is a subtype of int in Python, so any
# isinstance check below tests bool before int.
ConditionalValue = Union[bool, int, float, str, ConditionalSpecialValue]


class ConditionalActivity(str, enum.Enum):
    """Whether a source span is compiled, skipped, or undecidable."""

    ACTIVE = "active"
    INACTIVE = "inactive"
    UNKNOWN = "unknown"


@dataclass(frozen=True, slots=True)
class ConditionalCompilationEnvironment:
    """Compiler and project #Const inputs for directive evaluation."""

    compiler_constants: Mapping[str, ConditionalValue] | None = None
    project_constants: Mapping[str, ConditionalValue] | None = None


@dataclass(frozen=True, slots=True)
class ConditionalContainer:
    """Where a directive occurs: module level, or inside a named procedure."""

    kind: str  # "module" | "procedure"
    name: str | None = None
    span: Span | None = None


@dataclass(frozen=True, slots=True)
class ConditionalDirectiveOccurrence:
    directive: ConditionalDirectiveNode
    container: ConditionalContainer


@dataclass(frozen=True, slots=True)
class ConditionalConstDefinition:
    name: str
    name_span: Span
    value: ConditionalValue | None
    directive: ConditionalDirectiveNode
    value_raw: str | None = None


@dataclass(frozen=True, slots=True)
class ConditionalCompilationIndex:
    directives: list[ConditionalDirectiveOccurrence]
    constants: list[ConditionalConstDefinition]


# The compiler constants of 64-bit Office on Windows. One that is on is 1, not
# True (-1): `#If Not Win64` is true there, since Not 1 is -2, and
# `#If VBA7 = True` is false (XLIDE issue #214, measured in 64-bit Excel 16.0).
# One that is off is 0.
DEFAULT_COMPILER_CONSTANTS: Mapping[str, ConditionalValue] = {
    "VBA7": 1,
    "VBA6": 1,
    "Win64": 1,
    # Win32 is on in 64-bit Office as well: it means Windows, not a width (XLIDE
    # issue #192, measured in 64-bit Excel 16.0). Win16 is off.
    "Win32": 1,
    "Win16": 0,
    "Mac": 0,
    # TWINBASIC is a twinBASIC compiler auto-constant; in Excel VBA it is undefined,
    # and an undefined name is 0 (VBE-oracle verified). Modern libraries gate
    # twinBASIC-only intrinsics behind #If TWINBASIC, so without this default those
    # inactive branches were analyzed and produced false positives. Its value in VBA
    # is known, so it is a default, not left unknown.
    "TWINBASIC": 0,
}


def _effective_environment(
    env: ConditionalCompilationEnvironment | None,
) -> ConditionalCompilationEnvironment:
    env = env if env is not None else ConditionalCompilationEnvironment()
    merged: dict[str, ConditionalValue] = dict(DEFAULT_COMPILER_CONSTANTS)
    merged.update(env.compiler_constants or {})
    return ConditionalCompilationEnvironment(
        compiler_constants=merged, project_constants=env.project_constants
    )


@dataclass(slots=True)
class _ConditionalFrame:
    parent: ConditionalActivity
    current: ConditionalActivity
    seen_true: bool
    seen_unknown: bool


@dataclass(frozen=True, slots=True)
class _ConditionalArm:
    """One arm of one `#If` chain, as a persistent stack.

    ``parent`` is the enclosing chain's arm. Immutable, so an event can keep the
    arm that was in effect when it was recorded without copying the stack, and
    one arm is one object, which makes identity the comparison for
    ``in_same_branch``.
    """

    chain: int
    #: 0 for the `#If`, then one per `#ElseIf` / `#Else`.
    index: int
    parent: "_ConditionalArm | None"


def _arms_diverge(a: _ConditionalArm | None, b: _ConditionalArm | None) -> bool:
    """Whether two arm stacks disagree about which arm of a shared chain they are
    in. The branches then exclude each other whatever the constants are worth.
    Stacks are as deep as the source nests directives, so the walk is short."""
    outer = a
    while outer is not None:
        inner = b
        while inner is not None:
            if outer.chain == inner.chain:
                return outer.index != inner.index
            inner = inner.parent
        outer = outer.parent
    return False


@dataclass(slots=True)
class _ConditionalActivityEvent:
    start: int
    activity: ConditionalActivity
    branch: _ConditionalArm | None = None


class ConditionalActivityTracker:
    """Per-span activity lookup over one forward directive sweep (binary search)."""

    __slots__ = ("_events",)

    def __init__(self, events: list[_ConditionalActivityEvent]) -> None:
        self._events = events

    def _event_for_span(self, span: Span) -> _ConditionalActivityEvent | None:
        # Directives starting at or after the queried offset are not applied.
        lo = -1
        hi = len(self._events) - 1
        while lo < hi:
            mid = (lo + hi + 1) >> 1
            if self._events[mid].start < span.start:
                lo = mid
            else:
                hi = mid - 1
        return self._events[lo] if lo >= 0 else None

    def activity_for_span(self, span: Span) -> ConditionalActivity:
        event = self._event_for_span(span)
        return event.activity if event is not None else ConditionalActivity.ACTIVE

    def is_inactive(self, span: Span) -> bool:
        return self.activity_for_span(span) is ConditionalActivity.INACTIVE

    def mutually_exclusive(self, a: Span, b: Span) -> bool:
        """Whether the two spans sit in different arms of one `#If` chain, and so
        are never compiled together however the constants evaluate."""
        left = self._event_for_span(a)
        right = self._event_for_span(b)
        return _arms_diverge(
            left.branch if left is not None else None,
            right.branch if right is not None else None,
        )

    def in_same_branch(self, a: Span, b: Span) -> bool:
        """Whether the two spans sit under exactly the same arms, so every build
        either compiles both or neither.

        Stricter than "not mutually exclusive": spans in two SEPARATE chains are
        neither exclusive nor in the same branch, because a build may take one and
        not the other. Rules that pair two pieces of one construct need this, since
        a pairing made across different chains is a guess about a build that may
        never exist.
        """
        left = self._event_for_span(a)
        right = self._event_for_span(b)
        return (left.branch if left is not None else None) is (
            right.branch if right is not None else None
        )


def inactive_node_skip(
    activity: ConditionalActivityTracker | None,
) -> Callable[[BodyNode], bool] | None:
    """The `skip` test for iter_body_nodes that leaves out the nodes in inactive
    `#If` arms, or None when every node is active."""
    if activity is None:
        return None
    is_inactive = activity.is_inactive
    return lambda node: is_inactive(node.span)


def create_conditional_activity_tracker(
    module: ModuleNode, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalActivityTracker | None:
    """Build a per-span activity tracker, or None when the module has no directives."""
    if not module_has_conditional_directives(module):
        return None
    effective_env = _effective_environment(env)
    events = _collect_conditional_activity_events(module, effective_env)
    return ConditionalActivityTracker(events)


def _collect_conditional_activity_events(
    module: ModuleNode, effective_env: ConditionalCompilationEnvironment
) -> list[_ConditionalActivityEvent]:
    directives = collect_conditional_directives(module)
    project_constants = _lowercased(effective_env.project_constants)
    stack: list[_ConditionalFrame] = []
    current = ConditionalActivity.ACTIVE
    branch: _ConditionalArm | None = None
    chains = 0
    events: list[_ConditionalActivityEvent] = []
    for occ in directives:
        current = _apply_conditional_directive(occ.directive, effective_env, project_constants, stack, current)
        kind = occ.directive.directive_kind
        if kind is ConditionalDirectiveKind.IF:
            branch = _ConditionalArm(chain=chains, index=0, parent=branch)
            chains += 1
        elif kind in (ConditionalDirectiveKind.ELSE_IF, ConditionalDirectiveKind.ELSE):
            # An `#ElseIf` with no open `#If` is a parse-level error; leave the
            # stack alone rather than inventing an arm for it.
            if branch is not None:
                branch = _ConditionalArm(chain=branch.chain, index=branch.index + 1, parent=branch.parent)
        elif kind is ConditionalDirectiveKind.END_IF:
            branch = branch.parent if branch is not None else None
        events.append(
            _ConditionalActivityEvent(start=occ.directive.span.start, activity=current, branch=branch)
        )
    return events


def module_has_conditional_directives(module: ModuleNode) -> bool:
    """True when the module contains any #If/#Const directive at module level, inside a procedure body, or attached to an Enum/Type."""
    for member in module.members:
        if isinstance(member, ConditionalDirectiveNode):
            return True
        if isinstance(member, ProcedureNode) and _body_has_conditional_directives(member.body):
            return True
        if isinstance(member, (EnumNode, TypeNode)) and len(member.directives or []) > 0:
            return True
    return False


def index_conditional_compilation(
    module: ModuleNode, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalCompilationIndex:
    """Index a module's conditional compilation: all directive occurrences plus the resolved #Const definitions and their values."""
    directives = collect_conditional_directives(module)
    constants = _collect_conditional_constants(directives, _effective_environment(env))
    return ConditionalCompilationIndex(directives=directives, constants=constants)


def collect_conditional_directives(module: ModuleNode) -> list[ConditionalDirectiveOccurrence]:
    """Gather every conditional directive in the module (module level, procedure bodies, and Enum/Type members), sorted by source offset."""
    out: list[ConditionalDirectiveOccurrence] = []
    for member in module.members:
        if isinstance(member, ConditionalDirectiveNode):
            out.append(ConditionalDirectiveOccurrence(directive=member, container=ConditionalContainer(kind="module")))
        elif isinstance(member, ProcedureNode):
            _collect_body_directives(member.body, member, out)
        elif isinstance(member, (EnumNode, TypeNode)):
            for directive in member.directives or []:
                out.append(
                    ConditionalDirectiveOccurrence(directive=directive, container=ConditionalContainer(kind="module"))
                )
    out.sort(key=lambda o: o.directive.span.start)
    return out


_PROJECT_CONSTANT_INTEGER_RE = re.compile(r"[+-]?[0-9]+")


def _is_project_constant_name(name: str) -> bool:
    """XLIDE's ``/^\\p{L}[\\p{L}\\p{N}_]*$/u``, checked by Unicode category."""
    return (
        bool(name)
        and unicodedata.category(name[0]).startswith("L")
        and all(ch == "_" or unicodedata.category(ch)[0] in ("L", "N") for ch in name[1:])
    )


def parse_project_conditional_constants(raw: str | None) -> dict[str, ConditionalValue]:
    """The project's conditional compilation arguments, as Project Properties
    stores them (``"DEBUG_MODE = 1 : TRACE = 0"``, the PROJECTCONSTANTS record).

    Ported from parseProjectConditionalConstants in conditionalCompilation.ts:
    entries split on ``:``; a name that is not an identifier is skipped; an integer
    value is a number and anything else stays text. A name is a VBA identifier,
    whose letters are any the project's code page holds, so a name opening with
    an E acute is one (XLIDE issue #207).
    """
    constants: dict[str, ConditionalValue] = {}
    for entry in (raw or "").split(":"):
        name, eq, value_text = entry.partition("=")
        if not eq:
            continue
        name = js_trim(name)
        value_text = js_trim(value_text)
        if not _is_project_constant_name(name):
            continue
        if _PROJECT_CONSTANT_INTEGER_RE.fullmatch(value_text):
            constants[name] = _js_value(float(value_text))
        else:
            constants[name] = value_text
    return constants


def compiler_constants_with_defaults(
    env: ConditionalCompilationEnvironment | None = None,
) -> dict[str, ConditionalValue]:
    """The constants a directive sees, the defaults for 64-bit Office included:
    what conditional_compiler_constants gives for the environment the branch
    activity is decided in (XLIDE issue #215)."""
    return conditional_compiler_constants(_effective_environment(env))


def conditional_compiler_constants(
    env: ConditionalCompilationEnvironment | None = None,
) -> dict[str, ConditionalValue]:
    """Merge the compiler and project #Const values into one lowercased-name lookup, with project constants overriding compiler ones."""
    env = env if env is not None else ConditionalCompilationEnvironment()
    constants: dict[str, ConditionalValue] = {}
    for name, value in (env.compiler_constants or {}).items():
        constants[name.lower()] = value
    for name, value in (env.project_constants or {}).items():
        constants[name.lower()] = value
    return constants


def evaluate_conditional_expression(
    expression: str | None, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalValue | None:
    """Tokenize and evaluate a single #If/#Const expression against the environment's constants, returning its value or None when undecidable."""
    if expression is None or js_trim(expression) == "":
        return None
    parser = _ConditionalExpressionParser(
        _directive_expression_tokens(expression),
        conditional_compiler_constants(env),
        env is not None and env.project_constants is not None,
    )
    return parser.parse()


def _directive_expression_tokens(expression: str) -> list[VbaToken]:
    """A directive's expression as tokens. It is lexed after a throwaway ``x=``,
    so a ``#`` that opens it is a date literal and not the directive marker a
    ``#`` at the start of a statement is: ``#If #1/2/2000# > #1/1/2000# Then``
    (XLIDE issue #208)."""
    return [
        t
        for t in tokenize(f"x={expression}")[2:]
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]


def conditional_activity_at_offset(
    module: ModuleNode, offset: int, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalActivity:
    """Replay the directive stack up to a source offset and report whether that point is active, inactive, or unknown."""
    effective_env = _effective_environment(env)
    directives = collect_conditional_directives(module)
    project_constants = _lowercased(effective_env.project_constants)
    stack: list[_ConditionalFrame] = []
    current = ConditionalActivity.ACTIVE
    for occ in directives:
        if occ.directive.span.start >= offset:
            break
        current = _apply_conditional_directive(occ.directive, effective_env, project_constants, stack, current)
    return current


def conditional_activity_for_span(
    module: ModuleNode, span: Span, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalActivity:
    """Conditional activity (active/inactive/unknown) at the start of a source span."""
    return conditional_activity_at_offset(module, span.start, env)


def null_condition_directives(
    module: ModuleNode, env: ConditionalCompilationEnvironment | None = None
) -> list[ConditionalDirectiveNode]:
    """The `#If` and `#ElseIf` lines whose condition is Null, which the VBE
    refuses to compile: "Invalid use of Null" (XLIDE issue #208, Excel 16.0).
    `#If Null`, `#If Null = 1`, `#If Not Null` and a `#Const N = Null` read by
    `#If N` all are. Only a line the VBE is sure to evaluate is listed: one in
    code that is compiled, and for `#ElseIf`, after arms that were all False."""
    if not module_has_conditional_directives(module):
        return []
    effective_env = _effective_environment(env)
    project_constants = _lowercased(effective_env.project_constants)
    stack: list[_ConditionalFrame] = []
    current = ConditionalActivity.ACTIVE
    out: list[ConditionalDirectiveNode] = []
    for occ in collect_conditional_directives(module):
        directive = occ.directive
        frame = stack[-1] if stack else None
        if directive.directive_kind is ConditionalDirectiveKind.IF:
            evaluated = current is ConditionalActivity.ACTIVE
        else:
            evaluated = (
                directive.directive_kind is ConditionalDirectiveKind.ELSE_IF
                and frame is not None
                and frame.parent is ConditionalActivity.ACTIVE
                and not frame.seen_true
                and not frame.seen_unknown
            )
        if evaluated:
            value = _evaluate_with_project_constants(directive.condition_raw, effective_env, project_constants)
            if value is not None and _is_null(value):
                out.append(directive)
        current = _apply_conditional_directive(directive, effective_env, project_constants, stack, current)
    return out


def _apply_conditional_directive(
    directive: ConditionalDirectiveNode,
    env: ConditionalCompilationEnvironment,
    project_constants: dict[str, ConditionalValue],
    stack: list[_ConditionalFrame],
    current: ConditionalActivity,
) -> ConditionalActivity:
    kind = directive.directive_kind
    if kind is ConditionalDirectiveKind.CONST:
        # A #Const defines its constant even inside a #If False: the VBE reads
        # every #Const line (XLIDE issue #192, measured in Excel 16.0).
        if directive.name:
            value = _evaluate_with_project_constants(directive.value_raw, env, project_constants)
            if value is not None:
                project_constants[directive.name.lower()] = value
        return current
    if kind is ConditionalDirectiveKind.IF:
        condition = _condition_activity(directive, env, project_constants)
        frame = _ConditionalFrame(
            parent=current,
            current=_combine_activity(current, condition),
            seen_true=condition is ConditionalActivity.ACTIVE,
            seen_unknown=condition is ConditionalActivity.UNKNOWN,
        )
        stack.append(frame)
        return frame.current
    if kind is ConditionalDirectiveKind.ELSE_IF:
        if not stack:
            return current
        frame = stack[-1]
        condition = _condition_activity(directive, env, project_constants)
        if frame.seen_true:
            frame.current = ConditionalActivity.INACTIVE
        elif frame.seen_unknown and condition is not ConditionalActivity.INACTIVE:
            frame.current = _combine_activity(frame.parent, ConditionalActivity.UNKNOWN)
        else:
            frame.current = _combine_activity(frame.parent, condition)
        frame.seen_true = frame.seen_true or (condition is ConditionalActivity.ACTIVE)
        frame.seen_unknown = frame.seen_unknown or (condition is ConditionalActivity.UNKNOWN)
        return frame.current
    if kind is ConditionalDirectiveKind.ELSE:
        if not stack:
            return current
        frame = stack[-1]
        if frame.seen_true:
            frame.current = ConditionalActivity.INACTIVE
        elif frame.seen_unknown:
            frame.current = _combine_activity(frame.parent, ConditionalActivity.UNKNOWN)
        else:
            frame.current = frame.parent
        frame.seen_true = True
        return frame.current
    if kind is ConditionalDirectiveKind.END_IF:
        popped = stack.pop() if stack else None
        return popped.parent if popped is not None else current
    # Unknown
    return current


def _collect_body_directives(
    body: list[BodyNode], procedure: ProcedureNode, out: list[ConditionalDirectiveOccurrence]
) -> None:
    for node in iter_body_nodes(body):
        if isinstance(node, ConditionalDirectiveNode):
            out.append(
                ConditionalDirectiveOccurrence(
                    directive=node,
                    container=ConditionalContainer(kind="procedure", name=procedure.name, span=procedure.span),
                )
            )


def _body_has_conditional_directives(body: list[BodyNode]) -> bool:
    return any(isinstance(node, ConditionalDirectiveNode) for node in iter_body_nodes(body))


def _collect_conditional_constants(
    directives: list[ConditionalDirectiveOccurrence], env: ConditionalCompilationEnvironment
) -> list[ConditionalConstDefinition]:
    project_constants = _lowercased(env.project_constants)
    constants: list[ConditionalConstDefinition] = []
    for occ in directives:
        directive = occ.directive
        if (
            directive.directive_kind is not ConditionalDirectiveKind.CONST
            or not directive.name
            or directive.name_span is None
        ):
            continue
        # The index historically supplies a project table even when the caller
        # did not, so an absent name is zero on this path.
        value = _evaluate_with_project_constants(directive.value_raw, env, project_constants, True)
        if value is not None:
            project_constants[directive.name.lower()] = value
        constants.append(
            ConditionalConstDefinition(
                name=directive.name,
                name_span=directive.name_span,
                value_raw=directive.value_raw,
                value=value,
                directive=directive,
            )
        )
    return constants


def _condition_activity(
    directive: ConditionalDirectiveNode,
    env: ConditionalCompilationEnvironment,
    project_constants: Mapping[str, ConditionalValue],
) -> ConditionalActivity:
    value = _evaluate_with_project_constants(directive.condition_raw, env, project_constants)
    holds = None if value is None else _truthy(value)
    if holds is None:
        return ConditionalActivity.UNKNOWN
    return ConditionalActivity.ACTIVE if holds else ConditionalActivity.INACTIVE


class _ConstantLookup:
    """The module's `#Const` values over the compiler constants, by lookup only:
    copying every preceding #Const value for every directive made a forward
    replay quadratic."""

    __slots__ = ("_project", "_compiler")

    def __init__(
        self, project: Mapping[str, ConditionalValue], compiler: Mapping[str, ConditionalValue]
    ) -> None:
        self._project = project
        self._compiler = compiler

    def get(self, name: str) -> ConditionalValue | None:
        if name in self._project:
            return self._project[name]
        return self._compiler.get(name)


def _evaluate_with_project_constants(
    expression: str | None,
    env: ConditionalCompilationEnvironment,
    project_constants: Mapping[str, ConditionalValue],
    undefined_is_zero: bool | None = None,
) -> ConditionalValue | None:
    if expression is None or js_trim(expression) == "":
        return None
    if undefined_is_zero is None:
        undefined_is_zero = env.project_constants is not None
    # The module's own `#Const` values ride in `project_constants` whether or not
    # the caller supplied the project's; only the caller's presence says an absent
    # name is provably undefined (XLIDE issue #102).
    compiler_constants = conditional_compiler_constants(
        ConditionalCompilationEnvironment(compiler_constants=env.compiler_constants)
    )
    return _ConditionalExpressionParser(
        _directive_expression_tokens(expression),
        _ConstantLookup(project_constants, compiler_constants),
        undefined_is_zero,
    ).parse()


def _combine_activity(
    parent: ConditionalActivity, condition: ConditionalActivity
) -> ConditionalActivity:
    if parent is ConditionalActivity.INACTIVE or condition is ConditionalActivity.INACTIVE:
        return ConditionalActivity.INACTIVE
    if parent is ConditionalActivity.UNKNOWN or condition is ConditionalActivity.UNKNOWN:
        return ConditionalActivity.UNKNOWN
    return ConditionalActivity.ACTIVE


_EMPTY = ConditionalSpecialValue("empty")
_NULL = ConditionalSpecialValue("null")
_NOTHING = ConditionalSpecialValue("nothing")


def _is_special(value: ConditionalValue, kind: str) -> bool:
    return isinstance(value, ConditionalSpecialValue) and value.kind == kind


def _is_null(value: ConditionalValue) -> bool:
    return _is_special(value, "null")


def _truthy(value: ConditionalValue) -> bool | None:
    """Whether a condition holds. None for Null, which the VBE refuses as a
    condition ("Invalid use of Null"), and for Nothing ("Invalid use of object")."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return len(value) > 0
    if value.kind == "empty":
        return False
    if value.kind == "date":
        return value.serial != 0
    return None


def _lowercased(
    source: Mapping[str, ConditionalValue] | None,
) -> dict[str, ConditionalValue]:
    out: dict[str, ConditionalValue] = {}
    for name, value in (source or {}).items():
        out[name.lower()] = value
    return out


_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
_D_EXPONENT_RE = re.compile(r"[dD]")

# Integers below this are exact as a JavaScript number.
_MAX_SAFE_INTEGER = 2**53 - 1


def _js_value(number: float) -> int | float:
    """A JavaScript number as the port holds one: a whole number in the safe
    range as int, so it prints and keys as JavaScript's String() does, and
    anything else as float."""
    if isinstance(number, int):
        number = float(number)
    if math.isfinite(number) and number.is_integer() and abs(number) <= _MAX_SAFE_INTEGER:
        return int(number)
    return number


def _to_int32(number: float) -> int:
    """JavaScript's ToInt32, which every bitwise operator applies to its operands."""
    if not math.isfinite(number):
        return 0
    wrapped = int(number) & 0xFFFFFFFF
    return wrapped - 0x100000000 if wrapped >= 0x80000000 else wrapped


class _ConstantGetter(Protocol):
    def get(self, name: str, /) -> ConditionalValue | None: ...


class _ConditionalExpressionParser:
    """Evaluates a #If or #Const expression as the VBE does (XLIDE issues #192
    and #208, measured in Excel 16.0). The operators and their order are VBA's
    own, loosest first: Imp, Eqv, Xor, Or, And, Not, the comparisons with Like
    and Is, &, + and -, Mod, \\, * and /, unary minus, ^. Not, And, Or, Xor, Eqv
    and Imp are bitwise on numbers, as in code: `Not 1` is -2, which is True,
    and `1 And 2` is 0. Two Booleans give a Boolean. Strings compare without
    regard to case, and so does Like: `"A" = "a"` and `"ABC" Like "a*"` are
    True. Hex and octal literals keep their width: &HFFFF is -1.

    Empty is 0 beside a number and "" beside a string. Null propagates through
    arithmetic and comparisons, is "" to &, and follows VBA's three-valued
    logic: `Null Or True` is True and `Null And False` is False. A date literal
    is its serial, so `#12:00:00 AM#` is False. `Nothing Is Nothing` is True.
    """

    __slots__ = ("_tokens", "_constants", "_undefined_is_zero", "_index")

    def __init__(self, tokens: list[VbaToken], constants: _ConstantGetter, undefined_is_zero: bool) -> None:
        self._tokens = tokens
        self._constants = constants
        # Whether a name no constant defines evaluates as the VBE evaluates it,
        # to 0. Not to Empty: `UNDEFINED & "x" = "x"` is False where
        # `Empty & "x" = "x"` is True (XLIDE issue #208). True only when the
        # caller supplied the project's own conditional constants, so an absent
        # name is provably undefined rather than unknown (XLIDE issue #102); a
        # module's `#Const` lines are folded into the same table before any `#If`
        # reads them.
        self._undefined_is_zero = undefined_is_zero
        self._index = 0

    def parse(self) -> ConditionalValue | None:
        try:
            value = self._parse_logical(0)
        except RecursionError:
            # Upstream recurses on the JavaScript stack, which nests far deeper
            # than Python's; an expression nested past it is left unknown.
            return None
        return value if self._index >= len(self._tokens) else None

    def _parse_logical(self, level: int) -> ConditionalValue | None:
        """Imp, Eqv, Xor, Or, And, loosest first; each level is left-associative."""
        if level == len(_LOGICAL_LEVELS):
            return self._parse_not()
        op = _LOGICAL_LEVELS[level]
        left = self._parse_logical(level + 1)
        while self._match_word(op):
            right = self._parse_logical(level + 1)
            left = None if left is None or right is None else _logical(op, left, right)
        return left

    def _parse_not(self) -> ConditionalValue | None:
        """`Not` binds looser than a comparison: `Not 1 = 2` is `Not (1 = 2)`."""
        if self._match_word("not"):
            value = self._parse_not()
            if isinstance(value, bool):
                return not value
            if value is not None and _is_null(value):
                return _NULL
            number = None if value is None else _whole_number(value)
            return None if number is None else ~_to_int32(number)
        return self._parse_comparison()

    def _parse_comparison(self) -> ConditionalValue | None:
        left = self._parse_concat()
        while True:
            relational = relational_operator_at(self._tokens, self._index)
            word = None if relational is not None else token_word(self._peek())
            if relational is None and word != "like" and word != "is":
                return left
            self._index += relational[1] if relational is not None else 1
            right = self._parse_concat()
            if left is None or right is None:
                left = None
            elif relational is not None:
                left = _compare(relational[0], left, right)
            else:
                left = _like(left, right) if word == "like" else _is(left, right)

    def _parse_concat(self) -> ConditionalValue | None:
        left = self._parse_additive()
        while self._peek_raw() == "&":
            self._index += 1
            right = self._parse_additive()
            left = None if left is None or right is None else _concat(left, right)
        return left

    def _parse_additive(self) -> ConditionalValue | None:
        left = self._parse_mod()
        while self._peek_raw() in ("+", "-"):
            op = self._tokens[self._index].raw_text
            self._index += 1
            right = self._parse_mod()
            if left is None or right is None:
                left = None
            elif op == "+" and isinstance(left, str) and isinstance(right, str):
                left = left + right
            else:
                left = _arithmetic(op, left, right)
        return left

    def _parse_mod(self) -> ConditionalValue | None:
        left = self._parse_integer_division()
        while self._match_word("mod"):
            right = self._parse_integer_division()
            left = None if left is None or right is None else _arithmetic("mod", left, right)
        return left

    def _parse_integer_division(self) -> ConditionalValue | None:
        left = self._parse_product()
        while self._peek_raw() == "\\":
            self._index += 1
            right = self._parse_product()
            left = None if left is None or right is None else _arithmetic("\\", left, right)
        return left

    def _parse_product(self) -> ConditionalValue | None:
        left = self._parse_negation()
        while self._peek_raw() in ("*", "/"):
            op = self._tokens[self._index].raw_text
            self._index += 1
            right = self._parse_negation()
            left = None if left is None or right is None else _arithmetic(op, left, right)
        return left

    def _parse_negation(self) -> ConditionalValue | None:
        """Unary minus binds looser than ^: `-2 ^ 2` is -4."""
        op = self._peek_raw()
        if op in ("-", "+"):
            self._index += 1
            return _signed(op, self._parse_negation())
        return self._parse_power()

    def _parse_power(self) -> ConditionalValue | None:
        left = self._parse_primary()
        while self._peek_raw() == "^":
            self._index += 1
            right = self._parse_negation_operand()
            left = None if left is None or right is None else _arithmetic("^", left, right)
        return left

    def _parse_negation_operand(self) -> ConditionalValue | None:
        """An exponent may carry its own sign: `2 ^ -1`."""
        op = self._peek_raw()
        if op in ("-", "+"):
            self._index += 1
            return _signed(op, self._parse_primary())
        return self._parse_primary()

    def _parse_primary(self) -> ConditionalValue | None:
        token = self._peek()
        if token is None:
            return None
        if token.raw_text == "(":
            self._index += 1
            value = self._parse_logical(0)
            if self._peek_raw() != ")":
                return None
            self._index += 1
            return value
        self._index += 1
        if token.kind is TokenKind.INTEGER_LITERAL:
            return parse_vba_integer_literal(token.raw_text)
        if token.kind is TokenKind.FLOAT_LITERAL:
            number = js_number(_D_EXPONENT_RE.sub("e", _FLOAT_SUFFIX_RE.sub("", token.raw_text, count=1), count=1))
            return _js_value(number) if math.isfinite(number) else None
        if token.kind is TokenKind.STRING_LITERAL:
            return token.raw_text[1:-1].replace('""', '"')
        if token.kind is TokenKind.DATE_LITERAL:
            serial = date_literal_serial(token.raw_text)
            return None if serial is None else ConditionalSpecialValue("date", serial)
        word = token_word(token)
        if word == "true":
            return True
        if word == "false":
            return False
        if word == "empty":
            return _EMPTY
        if word == "null":
            return _NULL
        if word == "nothing":
            return _NOTHING
        value = self._constants.get(word)
        if value is None and self._undefined_is_zero and token.kind is TokenKind.IDENTIFIER:
            return 0
        return value

    def _match_word(self, word: str) -> bool:
        if token_word(self._peek()) != word:
            return False
        self._index += 1
        return True

    def _peek(self) -> VbaToken | None:
        return self._tokens[self._index] if self._index < len(self._tokens) else None

    def _peek_raw(self) -> str | None:
        token = self._peek()
        return token.raw_text if token is not None else None


_LOGICAL_LEVELS: tuple[str, ...] = ("imp", "eqv", "xor", "or", "and")


def _number_of(value: ConditionalValue) -> int | float | None:
    """The number a value converts to: True is -1, Empty 0, a Date its serial, a
    numeric string its number."""
    if isinstance(value, bool):
        return -1 if value else 0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, ConditionalSpecialValue):
        if value.kind == "empty":
            return 0
        if value.kind == "date":
            return value.serial
        return None
    trimmed = js_trim(value)
    parsed = math.nan if len(trimmed) == 0 else js_number(trimmed)
    return _js_value(parsed) if math.isfinite(parsed) else None


def _whole_number(value: ConditionalValue) -> int | float | None:
    """A value as the whole number a bitwise operator reads."""
    number = _number_of(value)
    return None if number is None else _js_value(bankers_round(number))


def _text(value: ConditionalValue) -> str | None:
    """A value as the text & and Like read. Empty and Null are "". A Date's
    text is the locale's, and a number's is left alone where JavaScript would
    spell it otherwise than VBA (1E+20, 0.1 + 0.2), so neither is guessed."""
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (int, float)):
        number = _js_value(value)
        return str(number) if isinstance(number, int) else None
    if isinstance(value, str):
        return value
    return "" if value.kind in ("empty", "null") else None


def _concat(left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    """`&`: Null is "" beside anything but another Null."""
    if _is_null(left) and _is_null(right):
        return _NULL
    a = _text(left)
    b = _text(right)
    return None if a is None or b is None else a + b


def _signed(op: str, value: ConditionalValue | None) -> ConditionalValue | None:
    if value is None or _is_null(value):
        return value
    number = _number_of(value)
    if number is None:
        return None
    return _js_value(-number) if op == "-" else number


def _logical(op: str, left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    """The bitwise and Boolean operators, with Null as VBA treats it: unknown,
    so `Null And False` is False and `Null Or True` is True, and anything that
    depends on the Null is Null."""
    if _is_null(left) or _is_null(right):
        return _logical_with_null(op, left, right)
    if isinstance(left, bool) and isinstance(right, bool):
        if op == "and":
            return left and right
        if op == "or":
            return left or right
        if op == "xor":
            return left != right
        if op == "eqv":
            return left == right
        return (not left) or right
    a = _whole_number(left)
    b = _whole_number(right)
    if a is None or b is None:
        return None
    x = _to_int32(a)
    y = _to_int32(b)
    if op == "and":
        return x & y
    if op == "or":
        return x | y
    if op == "xor":
        return x ^ y
    if op == "eqv":
        return ~(x ^ y)
    return ~x | y


def _logical_with_null(op: str, left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    known = right if _is_null(left) else left
    if _is_null(known):
        return _NULL
    # Which value of the known side decides the result alone: every bit clear
    # for And, every bit set for Or. Imp is decided by a False left side or a
    # True right side.
    bits = (-1 if known else 0) if isinstance(known, bool) else _whole_number(known)
    if bits is None:
        return None

    def decided(result: int) -> ConditionalValue:
        return result != 0 if isinstance(known, bool) else result

    if op == "and":
        return decided(0) if bits == 0 else _NULL
    if op == "or":
        return decided(-1) if bits == -1 else _NULL
    if op == "imp":
        if known is left:
            return decided(-1) if bits == 0 else _NULL
        return decided(-1) if bits == -1 else _NULL
    return _NULL


def _arithmetic(op: str, left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    if _is_null(left) or _is_null(right):
        return _NULL
    a = _number_of(left)
    b = _number_of(right)
    if a is None or b is None:
        return None
    result = _numeric_result(op, a, b)
    # A Date plus or minus a number is a Date; two Dates subtracted are days.
    left_date = _is_special(left, "date")
    right_date = _is_special(right, "date")
    if result is not None and op in ("+", "-") and left_date != right_date and (left_date or op == "+"):
        return ConditionalSpecialValue("date", result)
    return result


def _numeric_result(op: str, a: float, b: float) -> int | float | None:
    # JavaScript numbers are doubles: compute in float, then hold a whole result
    # as int (_js_value).
    x = float(a)
    y = float(b)
    if op == "+":
        return _js_value(x + y)
    if op == "-":
        return _js_value(x - y)
    if op == "*":
        return _js_value(x * y)
    if op == "/":
        return None if y == 0 else _js_value(x / y)
    if op == "\\":
        divisor = bankers_round(y)
        if divisor == 0:
            return None
        quotient = bankers_round(x) / divisor
        return _js_value(math.trunc(quotient) if math.isfinite(quotient) else quotient)
    if op == "mod":
        divisor = bankers_round(y)
        if divisor == 0:
            return None
        dividend = bankers_round(x)
        if not math.isfinite(dividend) or math.isnan(divisor):
            return math.nan  # JavaScript's % of Infinity, or by NaN, is NaN
        return _js_value(math.fmod(dividend, divisor))
    # `^`: Math.pow, whose NaN and Infinity results are no value here.
    if math.isnan(y) or (math.isinf(y) and abs(x) == 1):
        return None
    try:
        power = math.pow(x, y)
    except (OverflowError, ValueError, ZeroDivisionError):
        return None
    return _js_value(power) if math.isfinite(power) else None


def _compare(op: str, left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    """A comparison: two strings compare as text, without regard to case, and
    so does a string against Empty, which is ""; anything else compares as
    numbers. Null against anything is Null."""
    if _is_null(left) or _is_null(right):
        return _NULL

    def as_text(value: ConditionalValue) -> str | None:
        if isinstance(value, str):
            return value
        return "" if _is_special(value, "empty") else None

    order: int
    text_left = as_text(left)
    text_right = as_text(right)
    if (
        text_left is not None
        and text_right is not None
        and (isinstance(left, str) or isinstance(right, str))
    ):
        a_text = text_left.lower()
        b_text = text_right.lower()
        order = -1 if a_text < b_text else 1 if a_text > b_text else 0
    else:
        a = _number_of(left)
        b = _number_of(right)
        if a is None or b is None:
            return None
        order = -1 if a < b else 1 if a > b else 0
    if op == "=":
        return order == 0
    if op == "<>":
        return order != 0
    if op == "<":
        return order < 0
    if op == ">":
        return order > 0
    if op == "<=":
        return order <= 0
    return order >= 0


def _like(left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    """`Like`, without regard to case: `"ABC" Like "a*"` and `"a" Like "[A-C]"` are True."""
    if _is_null(left) or _is_null(right):
        return _NULL
    subject = _text(left)
    pattern = _text(right)
    if subject is None or pattern is None:
        return None
    regex = _like_pattern_regex(pattern)
    return None if regex is None else regex.fullmatch(subject) is not None


def _like_pattern_regex(pattern: str) -> re.Pattern[str] | None:
    """A Like pattern as a regular expression: `?` one character, `*` any run,
    `#` a digit, and `[...]` a character list, which `!` negates and `a-z`
    spans. None for a pattern VBA refuses at run time, a `[` never closed."""
    source = ""
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "?":
            source += r"[\s\S]"
        elif ch == "*":
            source += r"[\s\S]*"
        elif ch == "#":
            source += "[0-9]"
        elif ch == "[":
            close = pattern.find("]", i + 1)
            if close < 0:
                return None
            items = pattern[i + 1 : close]
            negated = items.startswith("!")
            if negated:
                items = items[1:]
            if not items:
                # `[]` matches nothing at all, `[!]` any one character.
                source += r"[\s\S]" if negated else "(?!)"
            else:
                escaped = "".join(c if c == "-" else re.escape(c) for c in items)
                source += f"[^{escaped}]" if negated else f"[{escaped}]"
            i = close
        else:
            source += re.escape(ch)
        i += 1
    try:
        return re.compile(source, re.IGNORECASE)
    except re.error:
        # A range written backwards, `[z-a]`, is a run-time error in VBA too.
        return None


def _is(left: ConditionalValue, right: ConditionalValue) -> ConditionalValue | None:
    """`Is` compares object references, and the only one a directive can name is Nothing."""
    return True if _is_special(left, "nothing") and _is_special(right, "nothing") else None
