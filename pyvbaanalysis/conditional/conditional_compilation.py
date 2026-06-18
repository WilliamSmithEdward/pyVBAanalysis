"""Conditional-compilation activity tracking (MS-VBAL 3.4).

Ported from xlide_vscode/src/analyzer/conditional/conditionalCompilation.ts.
Replays the #If/#ElseIf/#Else/#End If directive stack with the default compiler
constants (VBA7=true, Win64=true, Win32=false, Mac=false) plus any project #Const
definitions, and reports whether a given source span is active, inactive, or
unknown. Unknown stays unknown: the analyzer never guesses a branch.
"""

from __future__ import annotations

import enum
import math
import re
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Union

from ..lexer.token_helpers import token_word
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
)

# A resolved conditional value. bool is a subtype of int in Python, so any
# isinstance check below tests bool before int.
ConditionalValue = Union[bool, int, float, str]


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


DEFAULT_COMPILER_CONSTANTS: Mapping[str, ConditionalValue] = {
    "VBA7": True,
    "Win64": True,
    "Win32": False,
    "Mac": False,
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


@dataclass(slots=True)
class _ConditionalActivityEvent:
    start: int
    activity: ConditionalActivity


class ConditionalActivityTracker:
    """Per-span activity lookup over one forward directive sweep (binary search)."""

    __slots__ = ("_events",)

    def __init__(self, events: list[_ConditionalActivityEvent]) -> None:
        self._events = events

    def activity_for_span(self, span: Span) -> ConditionalActivity:
        # Directives starting at or after the queried offset are not applied.
        lo = -1
        hi = len(self._events) - 1
        while lo < hi:
            mid = (lo + hi + 1) >> 1
            if self._events[mid].start < span.start:
                lo = mid
            else:
                hi = mid - 1
        return self._events[lo].activity if lo >= 0 else ConditionalActivity.ACTIVE

    def is_inactive(self, span: Span) -> bool:
        return self.activity_for_span(span) is ConditionalActivity.INACTIVE


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
    events: list[_ConditionalActivityEvent] = []
    for occ in directives:
        current = _apply_conditional_directive(occ.directive, effective_env, project_constants, stack, current)
        events.append(_ConditionalActivityEvent(start=occ.directive.span.start, activity=current))
    return events


def module_has_conditional_directives(module: ModuleNode) -> bool:
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
    directives = collect_conditional_directives(module)
    constants = _collect_conditional_constants(directives, _effective_environment(env))
    return ConditionalCompilationIndex(directives=directives, constants=constants)


def collect_conditional_directives(module: ModuleNode) -> list[ConditionalDirectiveOccurrence]:
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


def conditional_compiler_constants(
    env: ConditionalCompilationEnvironment | None = None,
) -> dict[str, ConditionalValue]:
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
    if expression is None or expression.strip() == "":
        return None
    tokens = [
        t
        for t in tokenize(expression)
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    parser = _ConditionalExpressionParser(tokens, conditional_compiler_constants(env))
    return parser.parse()


def conditional_activity_at_offset(
    module: ModuleNode, offset: int, env: ConditionalCompilationEnvironment | None = None
) -> ConditionalActivity:
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
    return conditional_activity_at_offset(module, span.start, env)


def _apply_conditional_directive(
    directive: ConditionalDirectiveNode,
    env: ConditionalCompilationEnvironment,
    project_constants: dict[str, ConditionalValue],
    stack: list[_ConditionalFrame],
    current: ConditionalActivity,
) -> ConditionalActivity:
    kind = directive.directive_kind
    if kind is ConditionalDirectiveKind.CONST:
        if current is ConditionalActivity.ACTIVE and directive.name:
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
    for node in body:
        if isinstance(node, ConditionalDirectiveNode):
            out.append(
                ConditionalDirectiveOccurrence(
                    directive=node,
                    container=ConditionalContainer(kind="procedure", name=procedure.name, span=procedure.span),
                )
            )
        else:
            child = getattr(node, "body", None)
            if isinstance(child, list):
                _collect_body_directives(child, procedure, out)


def _body_has_conditional_directives(body: list[BodyNode]) -> bool:
    for node in body:
        if isinstance(node, ConditionalDirectiveNode):
            return True
        child = getattr(node, "body", None)
        if isinstance(child, list) and _body_has_conditional_directives(child):
            return True
    return False


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
        value = evaluate_conditional_expression(
            directive.value_raw,
            ConditionalCompilationEnvironment(
                compiler_constants=env.compiler_constants, project_constants=dict(project_constants)
            ),
        )
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
    if value is None:
        return ConditionalActivity.UNKNOWN
    return ConditionalActivity.ACTIVE if _truthy(value) else ConditionalActivity.INACTIVE


def _evaluate_with_project_constants(
    expression: str | None,
    env: ConditionalCompilationEnvironment,
    project_constants: Mapping[str, ConditionalValue],
) -> ConditionalValue | None:
    return evaluate_conditional_expression(
        expression,
        ConditionalCompilationEnvironment(
            compiler_constants=env.compiler_constants, project_constants=dict(project_constants)
        ),
    )


def _combine_activity(
    parent: ConditionalActivity, condition: ConditionalActivity
) -> ConditionalActivity:
    if parent is ConditionalActivity.INACTIVE or condition is ConditionalActivity.INACTIVE:
        return ConditionalActivity.INACTIVE
    if parent is ConditionalActivity.UNKNOWN or condition is ConditionalActivity.UNKNOWN:
        return ConditionalActivity.UNKNOWN
    return ConditionalActivity.ACTIVE


def _truthy(value: ConditionalValue) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return len(value) > 0


def _lowercased(
    source: Mapping[str, ConditionalValue] | None,
) -> dict[str, ConditionalValue]:
    out: dict[str, ConditionalValue] = {}
    for name, value in (source or {}).items():
        out[name.lower()] = value
    return out


_NUM_SUFFIX_RE = re.compile(r"[!#@%&^]$")


def _js_number(raw: str) -> int | float | None:
    """Mirror JavaScript Number(): decimal/float/exponent parse, else None.

    VBA &H/&O literals are not in the JS numeric grammar, so they yield None
    (Number('&HFF') is NaN). Integer-valued results are returned as int so their
    string form has no decimal point, matching JS String() in comparisons.
    """
    text = _NUM_SUFFIX_RE.sub("", raw).strip()
    try:
        number = float(text)
    except ValueError:
        return None
    if not math.isfinite(number):
        return None
    return int(number) if number.is_integer() else number


def _normalized_comparison_value(value: ConditionalValue) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value).lower()


class _ConditionalExpressionParser:
    """Recursive-descent evaluator: Or > And > comparison > Not > primary."""

    __slots__ = ("_tokens", "_constants", "_index")

    def __init__(self, tokens: list[VbaToken], constants: Mapping[str, ConditionalValue]) -> None:
        self._tokens = tokens
        self._constants = constants
        self._index = 0

    def parse(self) -> ConditionalValue | None:
        value = self._parse_or()
        return value if self._index >= len(self._tokens) else None

    def _parse_or(self) -> ConditionalValue | None:
        left = self._parse_and()
        while self._match_word("or"):
            right = self._parse_and()
            if left is None or right is None:
                return None
            left = _truthy(left) or _truthy(right)
        return left

    def _parse_and(self) -> ConditionalValue | None:
        left = self._parse_comparison()
        while self._match_word("and"):
            right = self._parse_comparison()
            if left is None or right is None:
                return None
            left = _truthy(left) and _truthy(right)
        return left

    def _parse_comparison(self) -> ConditionalValue | None:
        left = self._parse_unary()
        op_token = self._peek()
        op = op_token.raw_text if op_token is not None else None
        if op != "=" and op != "<>":
            return left
        self._index += 1
        right = self._parse_unary()
        if left is None or right is None:
            return None
        same = _normalized_comparison_value(left) == _normalized_comparison_value(right)
        return same if op == "=" else not same

    def _parse_unary(self) -> ConditionalValue | None:
        if self._match_word("not"):
            value = self._parse_unary()
            return None if value is None else (not _truthy(value))
        return self._parse_primary()

    def _parse_primary(self) -> ConditionalValue | None:
        token = self._peek()
        if token is None:
            return None
        if token.raw_text == "(":
            self._index += 1
            value = self._parse_or()
            close = self._peek()
            if close is None or close.raw_text != ")":
                return None
            self._index += 1
            return value
        self._index += 1
        if token.kind is TokenKind.INTEGER_LITERAL or token.kind is TokenKind.FLOAT_LITERAL:
            return _js_number(token.raw_text)
        if token.kind is TokenKind.STRING_LITERAL:
            return token.raw_text[1:-1].replace('""', '"')
        word = token_word(token)
        if word == "true":
            return True
        if word == "false":
            return False
        return self._constants.get(word)

    def _match_word(self, word: str) -> bool:
        if token_word(self._peek()) != word:
            return False
        self._index += 1
        return True

    def _peek(self) -> VbaToken | None:
        return self._tokens[self._index] if self._index < len(self._tokens) else None
