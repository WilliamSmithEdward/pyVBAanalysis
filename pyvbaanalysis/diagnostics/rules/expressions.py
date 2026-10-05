"""Rule family: expression-syntax rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/expressions.ts: unbalanced
parentheses, division by a provably-zero divisor, invalid-expression-syntax
(incomplete member access, the unsupported `?` operator, invalid operator runs),
and the call-shape rules (the parenthesized/parenless Call-statement and
expression-call forms). The call-shape and member-access rules ride the
member-completion context and runtime-function surfaces.
"""

from __future__ import annotations

import json
import math
import re
from collections.abc import Callable, Container, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import (
    MemberCompletionContext,
    resolve_exact_member_completion,
)
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    evaluate_integer_constant_expression,
    resolve_raw_integer_constants,
)
from ...host.host_model import HostObjectModel
from ...js_compat import JS_WHITESPACE, js_number
from ...lexer.token_helpers import match_paren_from, relational_operator_at
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    IfBranchKind,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    SelectBlockNode,
    Span,
    StatementNode,
    WhileBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...runtime.vba_runtime import resolve_runtime_function, runtime_allows_explicit_call
from ...symbols.name_resolution import BareIdentifierContext
from ...symbols.symbol_model import (
    ModuleSymbols,
    VbaProcedureSignature,
    VbaSymbol,
    VbaSymbolKind,
    qualified_procedure_key,
)
from ...types.type_inference import (
    SourceDeclaredType,
    declared_type_for_source_binding,
    defaulted_straight_line,
    known_local_literal_values_at,
    procedure_symbol_for,
    string_constants_in_scope,
    type_environment_for,
    unreachable_statements_in,
)
from ...types.type_names import (
    is_known_scalar_type,
    is_numeric_type,
    is_provably_non_numeric_string,
    normalize_type,
)
from ...call.call_context import (
    explicit_call_statement_argument_without_parens,
    explicit_call_statement_target,
    standalone_empty_parenthesized_call_statement,
    standalone_multi_arg_parenthesized_call_statement,
)
from ..call_extraction import CallableTypeSignature, callable_accepts_zero_arguments, string_literal_value
from ..callable_signatures import (
    SourceNameScope,
    bare_callable_source_shadowed,
    callable_signature_for,
    callable_type_signatures_for,
    procedure_integer_constant_lookup,
    runtime_callable_source_shadowed,
    source_name_scope_for,
)
from ..argument_inference import nonnumeric_string_arithmetic_operand
from ..condition_value import ConditionFacts, condition_value
from ..const_expr import (
    collect_module_literal_integer_constants,
    fold_integer_expression_tokens,
)
from ..context import PushFn, statement_tokens
from ..function_results import FunctionResult, function_integer_result, known_function_results
from ..known_locals import KnownLocalValue
from ..loop_counters import StatementCounters, check_each_counter_pass, loop_counters_at
from ..straight_line_values import EMPTY_COLLECTION, straight_line_assignments
from ..string_conversion import is_invalid_boolean_string, is_invalid_date_string
from ..type_fields import module_types
from ..type_member_state import MemberState, MemberStatesAt, is_known_number, type_member_states_at
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    bare_assignment_target,
    block_footer_line_span,
    block_header_line_span,
    first_executable_token_index,
    raw_expression_tokens,
    token_name,
    token_text,
    top_level_operator_index,
)
from .arrays import (
    ElementOperand,
    FixedArrayBound,
    element_operand_ending_at,
    element_operand_starting_at,
    elements_written_in,
    known_array_shapes_at,
    module_option_base,
)
from .shared import is_bare_or_vba_qualified_intrinsic_call


def check_unbalanced_parens(
    source: str, push: PushFn, activity: ConditionalActivityTracker | None = None
) -> None:
    """Every parenthesis must be matched within its logical statement (a `(` left
    open at a statement boundary, or a stray `)`, is a VBE Syntax error)."""
    # Text under an inactive `#If` arm is never compiled, and `#If False Then` is a
    # common place to park notes (XLIDE issue #102).
    toks = (
        [tok for tok in tokenize_cached(source) if not activity.is_inactive(Span(tok.start, tok.end))]
        if activity is not None
        else tokenize_cached(source)
    )
    depth = 0
    open_offsets: list[int] = []
    flagged = False

    def flush() -> None:
        nonlocal depth, flagged
        if not flagged and depth > 0:
            off = open_offsets[0]
            push("unbalancedParens", "Unbalanced parentheses: a ')' is missing.", Span(off, off + 1))
        depth = 0
        open_offsets.clear()
        flagged = False

    for tok in toks:
        if tok.kind is TokenKind.NEWLINE:
            flush()
            continue
        if tok.kind is TokenKind.COLON and depth == 0:
            flush()
            continue
        if tok.kind is not TokenKind.PUNCTUATION:
            continue
        if tok.raw_text == "(":
            depth += 1
            open_offsets.append(tok.start)
        elif tok.raw_text == ")":
            if depth == 0:
                if not flagged:
                    push(
                        "unbalancedParens",
                        "Unbalanced parentheses: an unexpected ')' was found.",
                        Span(tok.start, tok.end),
                    )
                    flagged = True
            else:
                depth -= 1
                open_offsets.pop()
    flush()


_TYPE_SUFFIX = re.compile(r"[!#@%&^]$")
_D_EXPONENT = re.compile(r"[dD]")
_HEX = re.compile(r"^&[hH]([0-9A-Fa-f]+)$")
_OCTAL = re.compile(r"^&[oO]([0-7]+)$")
_FLOAT = re.compile(r"^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")


def check_division_by_zero_expressions(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    host_model: HostObjectModel | None = None,
) -> ProcedureStatementVisitor:
    """`/`, `\\`, or `Mod` against a provably-zero divisor raises Run-time error 11."""
    project_constants = resolve_raw_integer_constants(project_integer_constants or {}, {})
    module_constants = collect_module_literal_integer_constants(mod, activity, project_constants)
    types = module_types(source, mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        constants = procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        # A numeric member of a Type local is 0 until an assignment stores
        # another literal (XLIDE issue #253): `1 / t.a` raises 11.
        members_at: MemberStatesAt | None = (
            None
            if len(types) == 0
            else type_member_states_at(
                source, symbols, member, types, activity, module_option_base(mod, activity)
            )
        )
        state = _DivisionState(
            source=source,
            member=member,
            symbols=symbols,
            activity=activity,
            constants=constants,
            # A Function of the module that returns 0: `10 / F()` (XLIDE issue #448).
            results=known_function_results(source, mod, activity),
        )
        # A local the procedure never assigns is 0, and one whose every
        # assignment is `d = 0` is 0 too (XLIDE issue #119): `10 / d` raises 11.
        # So is one the last assignment before the division sets to 0, though
        # a later one changes it (XLIDE issue #180).
        values_at: _LocalValuesAt = known_local_literal_values_at(source, member, symbols, activity)
        # A loop counter on its first and last passes: `For i = 0 To 3` then
        # `1 / i` divides by 0 on the first (XLIDE issue #263).
        counters: StatementCounters = loop_counters_at(source, member.body, activity)
        lookup = _DivisionByZeroLookup(state)
        guards = _division_guard_ranges(member.body, activity)
        # A statement a known guard keeps from running (XLIDE issue #273).
        unreachable: Container[int] = unreachable_statements_in(source, member, symbols, activity)

        # `d = 0.4` then `10 Mod d` (XLIDE issue #239).
        def fraction_of(lower: str) -> int | float | None:
            local = state.known.get(lower)
            if local is not None and local.kind == "number" and constants.get(lower) is None:
                value = local.value
                return value if isinstance(value, (int, float)) else None
            return None

        # A dividend that is Null, or a Variant the straight line just set to
        # Null: Null divided by zero is Null and raises nothing (XLIDE issue #282,
        # measured in Excel 16.0).
        def null_at(stmt: LeafStatementNode) -> Callable[[VbaToken | None], bool]:
            def is_null(tok: VbaToken | None) -> bool:
                if token_text(tok) == "null":
                    return True
                held = state.held(stmt, tok)
                return held is not None and len(held) == 1 and token_text(held[0]) == "null"

            return is_null

        # A number past the Long range, written or just stored: a literal, or a
        # CDec of one (XLIDE issue #502).
        def past_long_at(stmt: LeafStatementNode) -> Callable[[VbaToken | None], bool]:
            def past_long(tok: VbaToken | None) -> bool:
                if _outside_long(tok):
                    return True
                held = state.held(stmt, tok)
                conversion = (
                    held is not None
                    and len(held) == 4
                    and token_text(held[0]) == "cdec"
                    and held[1].raw_text == "("
                    and held[3].raw_text == ")"
                )
                return held is not None and (
                    (len(held) == 1 and _outside_long(held[0])) or (conversion and _outside_long(held[2]))
                )

            return past_long

        def visitor(stmt: LeafStatementNode) -> None:
            if id(stmt) in unreachable:
                return
            state.known = values_at(stmt)
            state.stmt = stmt
            # A single-line If's branch sees the members less what its condition names.
            state.members = members_at(stmt, stmt.span.end) if members_at is not None else state.members

            def check(values: Mapping[str, int | float], report: PushFn) -> None:
                state.pass_values = values
                for message, span in _division_by_zero_divisors(
                    source, stmt.span, lookup, guards, fraction_of, null_at(stmt), past_long_at(stmt)
                ):
                    report("divisionByZero", message, span)

            check_each_counter_pass(
                source, stmt.span, counters.get(stmt), lambda _atom, _counter: None, check, push
            )
            state.pass_values = {}

        return visitor

    return factory


# The values the rules read at a statement: knownLocalLiteralValuesAt's
# function, and straightLineAssignments' reaching values.
_LocalValuesAt = Callable[[BodyNode | None], Mapping[str, KnownLocalValue]]
# Upstream's Maps keyed by node: the port keys node maps by id().
_Reaching = Mapping[int, Mapping[str, Sequence[VbaToken]]]

# `"0"`, `" -00.0 "`, `".0"`: a string that divides as 0 (XLIDE issue #491);
# `\s` as JavaScript reads it.
_ZERO_STRING_RE = re.compile(r"[" + JS_WHITESPACE + r"]*[+-]?(?:0+\.?0*|\.0+)[" + JS_WHITESPACE + r"]*")
_COUNT_MEMBER_RE = re.compile(r"[a-z_][a-z0-9_]*[.]count", re.IGNORECASE | re.ASCII)
_PAST_LONG_TEXT_RE = re.compile(
    r"[" + JS_WHITESPACE + r"]*[+-]?[0-9]+(?:\.[0-9]*)?(?:[eE][+-]?[0-9]+)?[" + JS_WHITESPACE + r"]*"
)
_NUMERIC_TYPE_SUFFIX_RE = re.compile(r"[!#@%&^]$")


def _js_is_integer(value: object) -> bool:
    """JavaScript's Number.isInteger."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return isinstance(value, int) or (math.isfinite(value) and value.is_integer())


def _outside_long(value: VbaToken | None) -> bool:
    """A literal (or a string literal's text) whose number rounds past the Long range."""
    if value is None:
        return False
    if value.kind is TokenKind.STRING_LITERAL:
        text: str | None = value.raw_text[1:-1]
    elif value.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        text = _NUMERIC_TYPE_SUFFIX_RE.sub("", value.raw_text)
    else:
        text = None
    number = js_number(text) if text is not None and _PAST_LONG_TEXT_RE.fullmatch(text) else math.nan
    # Math.round rounds .5 up.
    return math.isfinite(number) and abs(math.floor(number + 0.5)) > 2147483647


class _DivisionState:
    """The per-statement facts the division lookup reads (upstream's closure
    variables in checkDivisionByZeroExpressions)."""

    __slots__ = (
        "source",
        "member",
        "symbols",
        "activity",
        "constants",
        "results",
        "known",
        "members",
        "pass_values",
        "stmt",
        "_reaching",
        "_defaulted",
    )

    def __init__(
        self,
        *,
        source: str,
        member: ProcedureNode,
        symbols: ModuleSymbols,
        activity: ConditionalActivityTracker | None,
        constants: IntegerConstantLookup,
        results: Mapping[str, FunctionResult],
    ) -> None:
        self.source = source
        self.member = member
        self.symbols = symbols
        self.activity = activity
        self.constants = constants
        self.results = results
        self.known: Mapping[str, KnownLocalValue] = {}
        self.members: Mapping[str, MemberState] = {}
        self.pass_values: Mapping[str, int | float] = {}
        self.stmt: LeafStatementNode | None = None
        self._reaching: _Reaching | None = None
        self._defaulted: _Reaching | None = None

    def reaching(self) -> _Reaching:
        if self._reaching is None:
            self._reaching = straight_line_assignments(self.source, self.member.body, self.activity)
        return self._reaching

    def held(self, stmt: LeafStatementNode, tok: VbaToken | None) -> list[VbaToken] | None:
        """What the straight line last stored into the name `tok` at `stmt`,
        comments dropped; None when it is not known."""
        name = token_name(tok)
        if not name:
            return None
        return self.held_named(stmt, name.lower())

    def held_named(self, stmt: LeafStatementNode, lower: str) -> list[VbaToken] | None:
        at = self.reaching().get(id(stmt))
        held = at.get(lower) if at is not None else None
        return None if held is None else [t for t in held if t.kind is not TokenKind.COMMENT]

    def empty_collection_now(self, lower: str) -> bool:
        if self.stmt is None:
            return False
        if self._defaulted is None:
            self._defaulted = defaulted_straight_line(self.source, self.member, self.symbols, self.activity)
        at = self._defaulted.get(id(self.stmt))
        return at is not None and at.get(lower) is EMPTY_COLLECTION


class _DivisionByZeroLookup:
    """The IntegerConstantLookup checkDivisionByZeroExpressions builds: a loop
    counter's value on the pass being checked, the procedure's constants, then
    what the text fixes a local, a Variant, a Type member or a Collection's
    Count to, then a module Function's known result."""

    __slots__ = ("_state",)

    def __init__(self, state: _DivisionState) -> None:
        self._state = state

    def get(self, name: str, /) -> float | None:
        state = self._state
        lower = name.lower()
        pass_value = state.pass_values.get(lower)
        if pass_value is not None:
            # A counter's pass value is any number, as upstream's lookup returns it.
            return int(pass_value) if _js_is_integer(pass_value) else pass_value
        constant = state.constants.get(name)
        if constant is not None:
            return constant
        local = state.known.get(lower)
        if local is not None and local.kind == "empty":
            return 0
        # "0" divides as 0 (XLIDE issue #491, measured in Excel 16.0).
        if (
            local is not None
            and local.kind == "string"
            and not local.content_mutated
            and _ZERO_STRING_RE.fullmatch(str(local.value)) is not None
        ):
            return 0
        # A Variant the straight line just set to Empty or CDec(0).
        held = state.held_named(state.stmt, lower) if state.stmt is not None else None
        if held is not None and (
            (len(held) == 1 and token_text(held[0]) == "empty")
            or _zero_conversion_call_end(held, 0) == len(held) - 1
        ):
            return 0
        field = state.members.get(lower)
        if is_known_number(field) and _js_is_integer(field.number):
            return int(field.number)
        # `c.Count` of a Collection nothing has added to (XLIDE issue #614,
        # measured in Excel 16.0: `10 / c.Count` raises 11).
        if _COUNT_MEMBER_RE.fullmatch(name) is not None and state.empty_collection_now(
            name[: -len(".count")].lower()
        ):
            return 0
        if local is not None and local.kind == "number" and _js_is_integer(local.value):
            return int(local.value)
        if local is not None:
            return None
        result = function_integer_result(name, state.results, state.member, state.symbols)
        if isinstance(result, float) and _js_is_integer(result):
            return int(result)
        # Upstream's lookup hands back any number; the evaluator reads it as it is.
        return result


def check_string_arithmetic_operands(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """An arithmetic operator raises error 13 for a nonnumeric string operand whatever
    the result goes into (XLIDE issue #119; each measured in Excel 16.0):
    `v = "abc" + 1` into a Variant, `Main = "abc" + 1` as a function result,
    `Not "abc"`, `-"abc"`, `If "abc" = 1 Then`, and `s * 2` with s holding "abc".
    The assignment and argument rules only saw the numeric-target case. `+` and the
    comparisons need a NUMBER on the other side, because two strings concatenate and
    compare as text; `- * / \\ ^ Mod` and the unary forms always coerce."""
    option_base = module_option_base(mod, activity)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        values_at: _LocalValuesAt = known_local_literal_values_at(source, member, symbols, activity)
        # A statement a known guard keeps from running (XLIDE issue #273).
        unreachable: Container[int] = unreachable_statements_in(source, member, symbols, activity)
        known: Mapping[str, KnownLocalValue] = {}
        # The arrays the locals hold at the statement, read only when an
        # operand indexes one: `v(1) + 1` with v = Array("1", "b") (XLIDE issue #260).
        shapes_at: Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]] | None = None
        current: LeafStatementNode | None = None
        shapes: Mapping[str, FixedArrayBound] | None = None
        written: AbstractSet[str] | None = None

        def shapes_here() -> Mapping[str, FixedArrayBound]:
            nonlocal shapes_at, shapes, written
            if current is None:
                return _EMPTY_SHAPES
            if shapes is None:
                shapes_of = shapes_at
                if shapes_of is None:
                    shapes_of = known_array_shapes_at(source, symbols, member, activity, option_base)
                    shapes_at = shapes_of
                written_now = written
                if written_now is None:
                    written_now = elements_written_in(source, member, activity)
                    written = written_now
                all_shapes = shapes_of(current)
                shapes = (
                    all_shapes
                    if len(written_now) == 0
                    else {lower: bound for lower, bound in all_shapes.items() if lower not in written_now}
                )
            return shapes

        def element_span(span_start: int, toks: Sequence[VbaToken], element: ElementOperand) -> Span:
            return Span(span_start + toks[element.first].start, span_start + toks[element.last].end)

        def element_string(
            span_start: int, toks: Sequence[VbaToken], element: ElementOperand | None
        ) -> str | None:
            if element is None or not isinstance(element.value, str) or not is_provably_non_numeric_string(
                element.value
            ):
                return None
            span = element_span(span_start, toks, element)
            return (
                f"'{source[span.start : span.end]}', which holds "
                f"{json.dumps(element.value, ensure_ascii=False)}"
            )

        string_consts: Mapping[str, str] = string_constants_in_scope(symbols, member)

        def constant_of(tok: VbaToken) -> str | None:
            name = token_name(tok)
            lower = name.lower() if name else None
            return string_consts.get(lower) if lower and lower not in known else None

        def nonnumeric_string(tok: VbaToken | None) -> str | None:
            if tok is None:
                return None
            if tok.kind is TokenKind.STRING_LITERAL:
                value = string_literal_value(tok.raw_text)
                return f"string literal {tok.raw_text}" if is_provably_non_numeric_string(value) else None
            name = token_name(tok)
            local = known.get(name.lower()) if name else None
            if (
                local is not None
                and local.kind == "string"
                and not local.content_mutated
                and is_provably_non_numeric_string(str(local.value))
            ):
                return f"'{tok.raw_text}', which holds {json.dumps(local.value, ensure_ascii=False)}"
            constant = constant_of(tok)
            return (
                f"constant '{tok.raw_text}', which is {json.dumps(constant, ensure_ascii=False)}"
                if constant is not None and is_provably_non_numeric_string(constant)
                else None
            )

        # A condition converts to Boolean: "True" and numbers run, "yes"
        # and " True " raise 13 (XLIDE issue #191).
        def non_boolean_string(tok: VbaToken | None) -> str | None:
            if tok is not None and tok.kind is TokenKind.STRING_LITERAL:
                return (
                    f"string literal {tok.raw_text}"
                    if is_invalid_boolean_string(string_literal_value(tok.raw_text))
                    else None
                )
            name = token_name(tok)
            local = known.get(name.lower()) if name else None
            if (
                tok is not None
                and local is not None
                and local.kind == "string"
                and not local.content_mutated
                and is_invalid_boolean_string(str(local.value))
            ):
                return f"'{tok.raw_text}', which holds {json.dumps(local.value, ensure_ascii=False)}"
            return None

        def numeric(tok: VbaToken | None) -> bool:
            if tok is None:
                return False
            # A date adds and compares as a number: #1/1/2000# + "abc" raises.
            if tok.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.DATE_LITERAL):
                return True
            name = token_name(tok)
            if not name:
                return False
            local = known.get(name.lower())
            if local is not None and local.kind == "number":
                return True
            type_ = normalize_type(env.get(name.lower()))
            return type_ is not None and is_numeric_type(type_)

        # A comparison converts only between typed operands: a Variant on
        # either side compares without converting, so `v = "abc"` with v
        # holding 5 is False and `v = 5` with v holding "abc" is False too
        # (XLIDE issue #268, measured in Excel 16.0). The number side is a
        # literal or a name declared a number type, and the string side a
        # literal, a String or a Const.
        def declared_type(tok: VbaToken | None) -> str | None:
            name = token_name(tok)
            return normalize_type(env.get(name.lower())) if name else None

        def typed_number(tok: VbaToken | None) -> bool:
            if tok is not None and tok.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
                return True
            type_ = declared_type(tok)
            return type_ is not None and is_numeric_type(type_)

        def typed_string(tok: VbaToken | None) -> bool:
            if tok is None or tok.kind is TokenKind.STRING_LITERAL:
                return tok is not None
            type_ = declared_type(tok)
            return type_ == "string" or (type_ is None and constant_of(tok) is not None)

        def typed_string_value(tok: VbaToken | None) -> tuple[str, str] | None:
            """The string a typed string operand holds, for a Date or Boolean
            comparison, as (value, what)."""
            if tok is None or not typed_string(tok):
                return None
            if tok.kind is TokenKind.STRING_LITERAL:
                return (string_literal_value(tok.raw_text), f"string literal {tok.raw_text}")
            name = token_name(tok)
            local = known.get(name.lower()) if name else None
            if local is not None and local.kind == "string" and not local.content_mutated:
                return (
                    str(local.value),
                    f"'{tok.raw_text}', which holds {json.dumps(local.value, ensure_ascii=False)}",
                )
            constant = constant_of(tok)
            return (
                None
                if constant is None
                else (constant, f"constant '{tok.raw_text}', which is {json.dumps(constant, ensure_ascii=False)}")
            )

        def unreadable_as(typed: VbaToken | None, other: VbaToken | None) -> tuple[str, str] | None:
            """A typed Date or Boolean compared with a string it cannot read:
            `d = "abc"`, `b = "yes"`, as (what, as)."""
            type_ = declared_type(typed)
            string = typed_string_value(other) if type_ in ("date", "boolean") else None
            if string is None:
                return None
            value, what = string
            invalid = is_invalid_date_string(value) if type_ == "date" else is_invalid_boolean_string(value)
            return (what, "a Date" if type_ == "date" else "a Boolean") if invalid else None

        def report_conversion(span: Span, what: str, into: str) -> None:
            push(
                "stringArithmeticCoercion",
                f"{_js_replace_first(into, 'WHAT', what)}. This will raise Run-time error '13': Type mismatch.",
                span,
            )

        # A condition that is one string: `If answer Then`, `Do While "abc"`.
        def check_condition(span_start: int, toks: Sequence[VbaToken], from_: int, to: int, keyword: str) -> None:
            if to - from_ != 1:
                return
            tok = _at(toks, from_)
            what = non_boolean_string(tok)
            if what and tok is not None:
                report_conversion(
                    absolute_span(Span(span_start, span_start), tok), what, f"'{keyword}' converts WHAT to Boolean"
                )

        # `For i = 1 To "abc"`: a bound converts to the counter's number.
        def check_for_bounds(span_start: int, toks: Sequence[VbaToken]) -> None:
            eq = _find_index(toks, lambda tok: tok.raw_text == "=")
            to = _find_index(toks, lambda tok: token_text(tok) == "to")
            step = _find_index(toks, lambda tok: token_text(tok) == "step")
            end = _condition_end(toks, 0)
            bounds: list[tuple[int, int, str]] = [
                (eq + 1, to, "start"),
                (to + 1, step if step > 0 else end, "end"),
                *([(step + 1, end, "step")] if step > 0 else []),
            ]
            for from_, until, which in bounds:
                tok = _at(toks, from_)
                what = nonnumeric_string(tok) if until - from_ == 1 and from_ > 0 else None
                if what and tok is not None:
                    report_conversion(
                        absolute_span(Span(span_start, span_start), tok),
                        what,
                        f"For converts WHAT to a number for its {which}",
                    )

        # `Select Case 1` then `Case "abc"`: each value is compared as a number.
        # The selector is a typed number, Date or Boolean: a Variant selector
        # compares without converting (XLIDE issue #268).
        def check_case_values(span: Span, selector: VbaToken) -> None:
            toks = statement_tokens(source, span)
            if token_text(_at(toks, 0)) != "case" or token_text(_at(toks, 1)) == "else":
                return
            end = _condition_end(toks, 1)

            def judge(value: VbaToken | None) -> bool:
                if value is None:
                    return False
                unreadable = unreadable_as(selector, value)
                if unreadable is not None:
                    what, as_ = unreadable
                    report_conversion(
                        absolute_span(span, value), what, f"Case compares WHAT with {as_}, which cannot read it"
                    )
                    return True
                # A Case value converts to the selector's type even from a
                # Variant: `Select Case n` on a Long with `Case w`, w holding
                # "abc", raises 13 (measured), where `w = n` runs.
                what_ = nonnumeric_string(value) if typed_number(selector) else None
                if what_:
                    report_conversion(absolute_span(span, value), what_, "Case compares WHAT with a number")
                return what_ is not None

            from_ = 1
            for k in range(1, end + 1):
                if k == end or toks[k].raw_text == ",":
                    if k - from_ == 1:
                        judge(toks[from_])
                    elif (
                        k - from_ == 3
                        and token_text(toks[from_]) == "is"
                        and toks[from_ + 1].kind is TokenKind.OPERATOR
                    ):
                        # `Case Is > "abc"` compares the same way (XLIDE issue #243).
                        judge(toks[from_ + 2])
                    elif k - from_ == 3 and token_text(toks[from_ + 1]) == "to":
                        # `Case "a" To "z"`: the first end that fails (XLIDE issue #268).
                        if not judge(toks[from_]):
                            judge(toks[from_ + 2])
                    from_ = k + 1

        def scan_operators(
            span_start: int,
            toks: Sequence[VbaToken],
            assign_index: int,
            branch_assigns: AbstractSet[int] = frozenset(),
        ) -> None:
            def at(tok: VbaToken) -> Span:
                return absolute_span(Span(span_start, span_start), tok)

            for i, tok in enumerate(toks):
                if i == assign_index or i in branch_assigns:
                    continue
                word = token_text(tok)
                left = _at(toks, i - 1)
                right = _at(toks, i + 1)
                operator_text = tok.raw_text

                def report(span: Span, what: str, operator_text: str = operator_text) -> None:
                    push(
                        "stringArithmeticCoercion",
                        f"Operator '{operator_text}' coerces {what} to a number. This will raise "
                        "Run-time error '13': Type mismatch.",
                        span,
                    )

                if word in _LOGICAL_OPERATORS and tok.kind is TokenKind.KEYWORD:
                    left_text = (
                        nonnumeric_string(left)
                        if i - 1 != assign_index
                        and (i - 1) not in branch_assigns
                        and (i - 2 == assign_index or (i - 2) in branch_assigns or _stands_alone(toks, i - 2))
                        else None
                    )
                    right_text = nonnumeric_string(right) if _stands_alone(toks, i + 2) else None
                    if left_text and left is not None:
                        report(at(left), left_text)
                    elif right_text and right is not None:
                        report(at(right), right_text)
                    continue
                is_binary = (
                    tok.raw_text in _ARITHMETIC_OR_COMPARISON if tok.kind is TokenKind.OPERATOR else word == "mod"
                )
                # A keyword ends an operand unless the statement's words start
                # one there: `If Not s` (XLIDE issue #268).
                left_ends_operand = left is not None and (
                    left.kind is TokenKind.IDENTIFIER
                    or (left.kind is TokenKind.KEYWORD and token_text(left) not in _OPERAND_STARTING_KEYWORDS)
                    or left.kind in _OPERAND_END_LITERAL_KINDS
                    or left.raw_text == ")"
                )
                # An element operand: `v(1)` or `Split("1 b")(1)` (XLIDE issue #260).
                right_element: ElementOperand | None = (
                    element_operand_starting_at(toks, i + 1, shapes_here(), option_base)
                    if token_name(right) is not None and _raw_at(toks, i + 2) == "("
                    else None
                )

                def right_operand(
                    right: VbaToken | None = right, right_element: ElementOperand | None = right_element
                ) -> tuple[Span, str] | None:
                    what = nonnumeric_string(right)
                    if what is None:
                        what = element_string(span_start, toks, right_element)
                    if not what or right is None:
                        return None
                    return (
                        element_span(span_start, toks, right_element) if right_element is not None else at(right),
                        what,
                    )

                if word == "not" and not left_ends_operand:
                    # Not binds below the comparisons: `Not s = "y"` is
                    # Not (s = "y"), a Boolean (XLIDE issue #361).
                    if _not_operand_compares(toks, i + 1):
                        continue
                    # `Not (v)` coerces v as well.
                    inner = (
                        _at(toks, i + 2)
                        if right is not None and right.raw_text == "(" and _raw_at(toks, i + 3) == ")"
                        else None
                    )
                    inner_what = nonnumeric_string(inner) if inner is not None else None
                    operand = (at(inner), inner_what) if inner is not None and inner_what else right_operand()
                    if operand is not None:
                        report(*operand)
                    continue
                if not is_binary:
                    continue
                if not left_ends_operand or left is None:
                    # Unary `-"abc"` (a leading `+` too).
                    if tok.raw_text in ("-", "+"):
                        operand = right_operand()
                        if operand is not None:
                            report(*operand)
                    continue
                left_element: ElementOperand | None = (
                    element_operand_ending_at(toks, i - 1, shapes_here(), option_base)
                    if left.raw_text == ")"
                    else None
                )
                always_coerces = tok.raw_text in _ALWAYS_COERCING or word == "mod"
                left_what = nonnumeric_string(left)
                if left_what is None:
                    left_what = element_string(span_start, toks, left_element)
                left_string = (
                    (element_span(span_start, toks, left_element) if left_element is not None else at(left), left_what)
                    if left_what
                    else None
                )
                right_string = right_operand()
                if always_coerces:
                    if left_string is not None:
                        report(*left_string)
                    elif right_string is not None:
                        report(*right_string)
                    continue
                if tok.raw_text != "+":
                    # A comparison: typed operands only (XLIDE issue #268).
                    if (
                        left_string is not None
                        and left_element is None
                        and typed_string(left)
                        and typed_number(right)
                        and right_element is None
                    ):
                        report(*left_string)
                    elif (
                        right_string is not None
                        and right_element is None
                        and typed_string(right)
                        and typed_number(left)
                        and left_element is None
                    ):
                        report(*right_string)
                    else:
                        from_left = unreadable_as(left, right)
                        unreadable = from_left if from_left is not None else unreadable_as(right, left)
                        if unreadable is not None:
                            what, as_ = unreadable
                            target = right if from_left is not None and right is not None else left
                            push(
                                "stringArithmeticCoercion",
                                f"Operator '{tok.raw_text}' compares {what} with {as_}, which cannot read it. "
                                "This will raise Run-time error '13': Type mismatch.",
                                at(target),
                            )
                    continue
                # `+`: a string against a NUMBER, True and False included:
                # `s + True` with s = "True" raises 13, and `"1" + True` is 0
                # (XLIDE issue #331, measured in Excel 16.0).
                if left_string is not None and (
                    numeric(right) or _is_truth(right) or _is_number_value(right_element)
                ):
                    report(*left_string)
                elif right_string is not None and (
                    numeric(left) or _is_truth(left) or _is_number_value(left_element)
                ):
                    report(*right_string)
                else:
                    # `s + d`, `s + b`: a String a Date or Boolean cannot read
                    # (XLIDE issue #331, measured in Excel 16.0).
                    from_left = unreadable_as(left, right)
                    unreadable = from_left if from_left is not None else unreadable_as(right, left)
                    if unreadable is not None:
                        what, as_ = unreadable
                        target = right if from_left is not None and right is not None else left
                        push(
                            "stringArithmeticCoercion",
                            f"Operator '+' adds {what} to {as_}, which cannot read it. This will raise "
                            "Run-time error '13': Type mismatch.",
                            at(target),
                        )

        # Block headers and footers are no statements of their own, so the
        # walk reads them here: If, Do and Loop conditions, While, a For
        # loop's bounds and Case values against a number (XLIDE issue #191).
        known = values_at(None)
        # Only a string literal, or a local known to hold a string, can be
        # reported; a header with neither is not read.
        knows_strings = any(value.kind == "string" for value in known.values())

        def may_hold_string(span: Span) -> bool:
            return knows_strings or '"' in source[span.start : span.end]

        for_reaching: _Reaching | None = None
        # Upstream's visitBlocks recurses per block; this walks the same
        # pre-order on an explicit stack.
        stack: list[BodyNode] = list(reversed(member.body))
        while stack:
            node = stack.pop()
            children = getattr(node, "body", None)
            if (
                (activity is not None and activity.is_inactive(node.span))
                or id(node) in unreachable
                or not isinstance(children, list)
            ):
                continue
            # `For i = 1 To v` with v a Variant the straight line just set to
            # Null: a bound must be a number (XLIDE issue #332, measured: 94).
            if isinstance(node, ForBlockNode) and not node.each:
                header = block_header_line_span(source, node.span)
                head_toks = statement_tokens(source, header)
                to = _find_index(head_toks, lambda tok: token_text(tok) == "to")
                step = _find_index(head_toks, lambda tok: token_text(tok) == "step")
                eq = _find_index(head_toks, lambda tok: tok.raw_text == "=")
                for from_, until in (
                    (eq + 1, to),
                    (to + 1, step if step > 0 else len(head_toks)),
                    (step + 1, len(head_toks) if step > 0 else -1),
                ):
                    bound = _at(head_toks, from_)
                    name = token_name(bound) if until - from_ == 1 and from_ > 0 else None
                    if not name or bound is None:
                        continue
                    if for_reaching is None:
                        for_reaching = straight_line_assignments(source, member.body, activity)
                    reaching_here = for_reaching.get(id(node))
                    held_raw = reaching_here.get(name.lower()) if reaching_here is not None else None
                    held = (
                        None if held_raw is None else [t for t in held_raw if t.kind is not TokenKind.COMMENT]
                    )
                    if held is not None and len(held) == 1 and token_text(held[0]) == "null":
                        push(
                            "variantValueMisuse",
                            f"'{bound.raw_text}' holds Null here, and For needs a number for its bounds. "
                            "This will raise Run-time error '94': Invalid use of Null.",
                            absolute_span(header, bound),
                        )
            # The opening line sees what reaches the block: `x = "abc"` then
            # `While x`, though the body assigns x again (XLIDE issue #424).
            known = values_at(node)
            reaching_string = any(value.kind == "string" for value in known.values())
            if (
                not isinstance(node, SelectBlockNode)
                and not reaching_string
                and not may_hold_string(block_header_line_span(source, node.span))
                and not (
                    isinstance(node, DoBlockNode) and may_hold_string(block_footer_line_span(source, node.span))
                )
            ):
                known = values_at(None)
                stack.extend(reversed(children))
                continue
            header = block_header_line_span(source, node.span)
            head_toks = statement_tokens(source, header)
            head = token_text(_at(head_toks, 0))
            scan_operators(header.start, head_toks, -1)
            if isinstance(node, IfBlockNode) and head == "if":
                check_condition(header.start, head_toks, 1, _condition_end(head_toks, 1), "If")
            elif isinstance(node, (DoBlockNode, WhileBlockNode)) and len(head_toks) > 1:
                keyword = token_text(head_toks[1 if head == "do" else 0])
                if keyword in ("while", "until"):
                    from_ = 2 if head == "do" else 1
                    check_condition(
                        header.start,
                        head_toks,
                        from_,
                        _condition_end(head_toks, from_),
                        "While" if keyword == "while" else "Until",
                    )
            if isinstance(node, DoBlockNode):
                footer = block_footer_line_span(source, node.span)
                foot_toks = statement_tokens(source, footer)
                loop_keyword = token_text(_at(foot_toks, 1))
                if token_text(_at(foot_toks, 0)) == "loop" and loop_keyword in ("while", "until"):
                    # The Loop line runs after the body, with what the body leaves.
                    entering = known
                    known = values_at(None)
                    scan_operators(footer.start, foot_toks, -1)
                    check_condition(
                        footer.start,
                        foot_toks,
                        2,
                        _condition_end(foot_toks, 2),
                        "While" if loop_keyword == "while" else "Until",
                    )
                    known = entering
            if (
                isinstance(node, ForBlockNode)
                and not node.each
                and node.control_variable
                and numeric(VbaToken(TokenKind.IDENTIFIER, node.control_variable, 0, 0, 0, 0))
            ):
                check_for_bounds(header.start, head_toks)
            # The selector is judged with each value: a typed number, Date or
            # Boolean converts; a Variant does not.
            if isinstance(node, SelectBlockNode) and head == "select" and _condition_end(head_toks, 2) == 3:
                for item in node.body:
                    if is_leaf_statement(item) and not (activity is not None and activity.is_inactive(item.span)):
                        check_case_values(item.span, head_toks[2])
            stack.extend(reversed(children))

        def visitor(stmt: LeafStatementNode) -> None:
            nonlocal known, current, shapes
            if id(stmt) in unreachable:
                return
            known = values_at(stmt)
            current = stmt
            shapes = None
            toks = statement_tokens(source, stmt.span)
            first = first_executable_token_index(toks)
            head = token_text(_at(toks, first))
            if head == "const":
                return
            # A one-line If and an ElseIf line are statements; so is IIf.
            if head in ("if", "elseif"):
                check_condition(
                    stmt.span.start, toks, first + 1, _condition_end(toks, first + 1), "If" if head == "if" else "ElseIf"
                )
            for k in range(len(toks) - 1):
                if (
                    len(toks[k].raw_text) == 3
                    and toks[k + 1].raw_text == "("
                    and token_text(toks[k]) == "iif"
                    and is_bare_or_vba_qualified_intrinsic_call(toks, k)
                ):
                    close = match_paren_from(toks, k + 1)
                    comma = _find_index(toks, lambda tok: tok.raw_text == ",", after=k + 1)
                    if close > 0 and comma > 0 and comma < close:
                        check_condition(stmt.span.start, toks, k + 2, comma, "IIf")
            # A string literal in arithmetic into a numeric variable is the
            # assignment rule's: it already names the target, and one report per
            # line is enough. A local holding the string is this rule's, since
            # the assignment rule reads only literals: `x = s + 1` into a Long
            # with s holding "abc" (XLIDE issue #180).
            bare = bare_assignment_target(source, stmt.span)
            if bare is not None:
                target_type = env.get(bare[0].lower())
                if target_type and nonnumeric_string_arithmetic_operand(target_type, bare[2], 0):
                    return
            # The assignment's own `=` stores; it compares nothing. So does the
            # `=` of an assignment in a one-line If's branch: `If L > 0 Then
            # tb = s` (XLIDE issue #342).
            branch_assigns: set[int] = set()
            branches = (stmt.single_line_if_branches or []) if isinstance(stmt, StatementNode) else []
            for branch in branches:
                if bare_assignment_target(source, branch) is not None:
                    at = next(
                        (
                            index
                            for index, tok in enumerate(toks)
                            if tok.raw_text == "=" and stmt.span.start + tok.start >= branch.start
                        ),
                        -1,
                    )
                    if at >= 0:
                        branch_assigns.add(at)
            scan_operators(
                stmt.span.start,
                toks,
                _find_index(toks, lambda tok: tok.raw_text == "=") if bare is not None else -1,
                branch_assigns,
            )

        return visitor

    return factory


# The operators that read both sides as numbers when either is a number.
_ARITHMETIC_OR_COMPARISON = frozenset({"+", "-", "*", "/", "\\", "^", "=", "<", ">", "<=", ">=", "<>"})
_ALWAYS_COERCING = frozenset({"-", "*", "/", "\\", "^"})
_OPERAND_END_LITERAL_KINDS = frozenset(
    {
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    }
)

# The operators that convert both operands to numbers and bind loosest.
_LOGICAL_OPERATORS = frozenset({"and", "or", "xor", "eqv", "imp"})

_EMPTY_SHAPES: Mapping[str, FixedArrayBound] = {}

_COMPARISON_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>"})

# Keywords after which an operand starts, so a `Not` or a sign there is unary.
_OPERAND_STARTING_KEYWORDS = frozenset(
    {
        "if", "elseif", "then", "else", "while", "until", "case", "to", "step", "and", "or", "xor",
        "eqv", "imp", "not", "mod", "like", "is", "call", "set", "let", "return",
    }
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """toks[i], or None where JavaScript reads undefined (a negative index too)."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _find_index(toks: Sequence[VbaToken], test: Callable[[VbaToken], bool], *, after: int = -1) -> int:
    """Array.prototype.findIndex over the tokens past `after`; -1 when none matches."""
    return next((index for index, tok in enumerate(toks) if index > after and test(tok)), -1)


def _condition_end(toks: Sequence[VbaToken], from_: int) -> int:
    """Where a condition starting at `from_` ends: its Then, else a comment,
    else the end of the statement."""
    then = _find_index(toks, lambda tok: token_text(tok) == "then", after=from_ - 1)
    comment = _find_index(toks, lambda tok: tok.kind is TokenKind.COMMENT, after=from_ - 1)
    return then if then >= 0 else comment if comment >= 0 else len(toks)


def _stands_alone(toks: Sequence[VbaToken], index: int) -> bool:
    """Logical operators convert a string operand to a number, and bind
    loosest: in `s = "yes" Or t` the string is the `=`'s. An operand is judged
    only when it stands alone between the operator and a boundary."""
    tok = _at(toks, index)
    if tok is None:
        return True
    word = token_text(tok)
    return (
        tok.raw_text in ("(", ")", ",", ":")
        or tok.kind is TokenKind.COMMENT
        or word in _LOGICAL_OPERATORS
        or word in ("then", "if", "elseif", "while", "until", "not")
    )


def _is_truth(operand: VbaToken | None) -> bool:
    return token_text(operand) in ("true", "false")


def _is_number_value(element: ElementOperand | None) -> bool:
    """typeof element?.value === 'number'."""
    return (
        element is not None
        and isinstance(element.value, (int, float))
        and not isinstance(element.value, bool)
    )


def _js_replace_first(text: str, pattern: str, replacement: str) -> str:
    """String.prototype.replace with a string pattern: the first occurrence only,
    with the replacement's `$$`, `$&`, `` $` `` and `$'` patterns expanded."""
    index = text.find(pattern)
    if index < 0:
        return text
    before = text[:index]
    after = text[index + len(pattern) :]
    out: list[str] = []
    i = 0
    while i < len(replacement):
        char = replacement[i]
        following = replacement[i + 1] if i + 1 < len(replacement) else ""
        if char == "$" and following in ("$", "&", "`", "'"):
            out.append({"$": "$", "&": pattern, "`": before, "'": after}[following])
            i += 2
            continue
        out.append(char)
        i += 1
    return before + "".join(out) + after


def _not_operand_compares(toks: Sequence[VbaToken], from_: int) -> bool:
    """Whether the operand of a Not starting at `from_` holds a comparison: Not
    takes everything up to the next And, Or, Xor, Eqv or Imp, a closing
    parenthesis, a comma or Then, and a comparison inside makes it a Boolean
    (XLIDE issue #361, measured in Excel 16.0)."""
    depth = 0
    for k in range(max(from_, 0), len(toks)):
        tok = toks[k]
        word = token_text(tok)
        if tok.raw_text == "(":
            depth += 1
        elif tok.raw_text == ")":
            if depth == 0:
                return False
            depth -= 1
        elif depth == 0:
            if word in _LOGICAL_OPERATORS or word == "then" or tok.raw_text in (",", ":"):
                return False
            if (
                tok.kind is TokenKind.OPERATOR and tok.raw_text in _COMPARISON_OPERATORS
            ) or word in ("like", "is"):
                return True
    return False


@dataclass(frozen=True, slots=True)
class _DivisionGuard:
    """A name a branch has tested non-zero, and the span the test covers."""

    name: str
    start: int
    end: int


def _division_guard_names(condition: Sequence[VbaToken]) -> tuple[set[str], set[str]]:
    """Which names an If condition proves non-zero on its Then arm (`SCALE_BY <> 0`,
    `n > 0`, `Not n = 0`, a bare `n`) and which it proves zero (`n = 0`, so the Else
    arm has the non-zero case), as (non-zero, zero). A constant that fails the test
    never reaches the division: `If SCALE_BY <> 0 Then x = 10 / SCALE_BY` with
    SCALE_BY = 0 runs clean (XLIDE issue #106, measured in Excel 16.0)."""
    non_zero: set[str] = set()
    zero: set[str] = set()
    words = [token_text(tok) for tok in condition]

    def name_at(index: int) -> str | None:
        name = token_name(condition[index]) if 0 <= index < len(condition) else None
        return name.lower() if name else None

    def is_zero(index: int) -> bool:
        tok = condition[index] if 0 <= index < len(condition) else None
        return (
            tok is not None
            and tok.kind is TokenKind.INTEGER_LITERAL
            and _ALL_ZEROS_RE.fullmatch(tok.raw_text) is not None
        )

    # Conjuncts each hold on the Then arm; a disjunction proves nothing.
    if "or" in words:
        return non_zero, zero
    start = 0
    for i in range(len(words) + 1):
        if i < len(words) and words[i] != "and":
            continue
        w = words[start:i]
        n0, n1, n2 = name_at(start), name_at(start + 1), name_at(start + 2)
        comparison = len(w) == 3 and w[1] in ("<>", ">", "<")
        if len(w) == 1 and n0:
            non_zero.add(n0)
        elif comparison and n0 and is_zero(start + 2):
            non_zero.add(n0)
        elif comparison and n2 and is_zero(start):
            non_zero.add(n2)
        elif len(w) == 3 and n0 and is_zero(start + 2) and w[1] == "=":
            zero.add(n0)
        elif len(w) == 4 and w[0] == "not" and n1 and w[2] == "=" and is_zero(start + 3):
            non_zero.add(n1)
        elif (
            len(w) == 6
            and w[0] == "not"
            and w[1] == "("
            and n2
            and w[3] == "="
            and is_zero(start + 4)
            and w[5] == ")"
        ):
            non_zero.add(n2)
        elif len(w) == 2 and w[0] == "not" and n1:
            zero.add(n1)
        start = i + 1
    return non_zero, zero


def _division_guard_ranges(
    body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> list[_DivisionGuard]:
    """The guards every block If in the body establishes for its arms."""
    out: list[_DivisionGuard] = []
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if not isinstance(node, IfBlockNode):
            continue
        for index, branch in enumerate(node.branches):
            if not branch.condition_raw:
                continue
            condition = raw_expression_tokens(branch.condition_raw)
            non_zero, zero = _division_guard_names(condition)
            for name in non_zero:
                out.append(_DivisionGuard(name, branch.span.start, branch.span.end))
            following = node.branches[index + 1] if index + 1 < len(node.branches) else None
            if following is not None and following.branch_kind is IfBranchKind.ELSE:
                for name in zero:
                    out.append(_DivisionGuard(name, following.span.start, following.span.end))
    return out


def _division_by_zero_divisors(
    source: str,
    span: Span,
    constants: IntegerConstantLookup,
    guards: Sequence[_DivisionGuard],
    fraction_of: Callable[[str], int | float | None] | None = None,
    is_null: Callable[[VbaToken | None], bool] = lambda _tok: False,
    past_long: Callable[[VbaToken | None], bool] = lambda _tok: False,
) -> list[tuple[str, Span]]:
    """(message, span) for each division whose divisor the text proves zero."""
    toks = statement_tokens(source, span)
    hits: list[tuple[str, Span]] = []
    # A single-line If guards its own arms.
    first = first_executable_token_index(toks)
    then_index = -1
    else_index = -1
    local_non_zero: set[str] = set()
    local_zero: set[str] = set()
    if token_text(toks[first] if first < len(toks) else None) == "if":
        then_index = _index_of_word_after(toks, first, "then")
        if then_index > 0:
            local_non_zero, local_zero = _division_guard_names(toks[first + 1 : then_index])
            else_index = _index_of_word_after(toks, then_index, "else")
    for i, tok in enumerate(toks):
        operator = _division_by_zero_operator_label(tok)
        if operator is None:
            continue
        divisor = _zero_divisor_token(source, span, toks, i + 1, constants)
        if divisor is None:
            divisor = _fractional_divisor_rounding_to_zero(toks, i + 1, operator, fraction_of, constants)
        if divisor is None:
            continue
        divisor_name = token_name(divisor[0]) if len(divisor) == 1 else None
        if divisor_name:
            lower = divisor_name.lower()
            in_else = else_index >= 0 and i > else_index
            in_then = then_index >= 0 and i > then_index and not in_else
            if (in_then and lower in local_non_zero) or (in_else and lower in local_zero):
                continue
            at = span.start + tok.start
            if any(guard.name == lower and guard.start <= at < guard.end for guard in guards):
                continue
        # `0 / 0` raises 6 (Overflow), not 11; `\` and `Mod` raise 11 for it (XLIDE
        # issue #106, measured in Excel 16.0).
        dividend = toks[i - 1] if i >= 1 else None
        if is_null(dividend) and _raw_at(toks, i - 2) != ".":
            continue
        # `\` and Mod convert the dividend to a Long first: one past the Long
        # range overflows (6) before the division (XLIDE issue #502).
        if operator != "/" and past_long(dividend) and _raw_at(toks, i - 2) != ".":
            continue
        dividend_name = token_name(dividend)
        dividend_zero = dividend is not None and (
            (
                dividend.kind is TokenKind.INTEGER_LITERAL
                and _ZERO_INTEGER_LITERAL_RE.fullmatch(dividend.raw_text) is not None
            )
            or (
                dividend.kind is TokenKind.FLOAT_LITERAL
                and _js_number_or_none(_FLOAT_TYPE_SUFFIX_RE.sub("", dividend.raw_text)) == 0
            )
            or (dividend_name is not None and constants.get(dividend_name.lower()) == 0)
        )
        message = (
            "Expression divides zero by zero with '/'. This will raise Run-time error '6': Overflow."
            if operator == "/" and dividend_zero
            else f"Expression uses '{operator}' with a zero divisor. This will raise Run-time "
            "error '11': Division by zero."
        )
        hits.append((message, _absolute_token_group_span(span, divisor)))
    return hits


_ALL_ZEROS_RE = re.compile(r"0+")
_ZERO_INTEGER_LITERAL_RE = re.compile(r"0+[%&^]?")
_FLOAT_TYPE_SUFFIX_RE = re.compile(r"[!#@]$")


def _index_of_word_after(toks: Sequence[VbaToken], after: int, word: str) -> int:
    """The index of the first token past `after` whose text is `word`, or -1."""
    return next((index for index, tok in enumerate(toks) if index > after and token_text(tok) == word), -1)


def _js_number_or_none(text: str) -> float | None:
    """JavaScript's Number() for a numeric literal's text, None where it is NaN."""
    try:
        return float(text)
    except ValueError:
        return None


def _fractional_divisor_rounding_to_zero(
    toks: Sequence[VbaToken],
    start: int,
    operator: str,
    fraction_of: Callable[[str], int | float | None] | None = None,
    constants: IntegerConstantLookup | None = None,
) -> list[VbaToken] | None:
    """`\\` and `Mod` round their operands to whole numbers first, with banker's
    rounding, so a literal divisor below 0.5 - or exactly 0.5 - is zero to them:
    `5 \\ 0.4` and `5 Mod 0.5` raise 11 (XLIDE issue #119, measured in Excel 16.0).
    So is a local holding such a value here: `d = 0.4` then `10 Mod d`
    (XLIDE issue #239)."""
    if operator == "/":
        return None
    name_raw = token_name(_at(toks, start))
    name = name_raw.lower() if name_raw else None
    held = (
        fraction_of(name)
        if name and fraction_of is not None and _is_divisor_atom_boundary(_at(toks, start + 1))
        else None
    )
    if held is not None:
        return [toks[start]] if held != 0 and abs(held) <= 0.5 else None
    # `10 \ Val("0.4")`: a call the folder reads to a fraction (XLIDE issue #703).
    if name and _raw_at(toks, start + 1) == "(" and _raw_at(toks, start - 1) != "." and constants is not None:
        end = match_paren_from(toks, start + 1)
        call = list(toks[start : end + 1]) if end > start else []
        value = (
            evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in call), constants)
            if len(call) > 0 and _is_divisor_atom_boundary(_at(toks, end + 1))
            else None
        )
        return call if value is not None and value != 0 and abs(value) <= 0.5 else None
    index = start
    group: list[VbaToken] = []
    signed = toks[index] if index < len(toks) else None
    if signed is not None and signed.kind is TokenKind.OPERATOR and signed.raw_text in ("-", "+"):
        group.append(signed)
        index += 1
    literal = toks[index] if index < len(toks) else None
    if (
        literal is None
        or literal.kind is not TokenKind.FLOAT_LITERAL
        or not _is_divisor_atom_boundary(toks[index + 1] if index + 1 < len(toks) else None)
    ):
        return None
    parsed = _js_number_or_none(_D_EXPONENT.sub("E", _FLOAT_TYPE_SUFFIX_RE.sub("", literal.raw_text)))
    if parsed is None or not math.isfinite(parsed) or abs(parsed) > 0.5:
        return None
    group.append(literal)
    return group


def _division_by_zero_operator_label(tok: VbaToken) -> str | None:
    text = token_text(tok)
    if text in ("/", "\\"):
        return text
    return "Mod" if text == "mod" else None


def _zero_divisor_token(
    source: str,
    span: Span,
    toks: list[VbaToken],
    start: int,
    constants: IntegerConstantLookup,
) -> list[VbaToken] | None:
    if start >= len(toks):
        return None
    first = toks[start]
    if first.raw_text == "(":
        close = match_paren_from(toks, start)
        if close < 0:
            return None
        return _zero_divisor_expression(source, span, toks, start + 1, close, constants)
    if first.kind is TokenKind.OPERATOR and first.raw_text in ("+", "-"):
        signed = _zero_divisor_atom_token_group(toks, start + 1, constants)
        return [first, *signed] if signed else None
    return _zero_divisor_atom_token_group(toks, start, constants)


def _zero_divisor_expression(
    source: str,
    span: Span,
    toks: list[VbaToken],
    start: int,
    end_exclusive: int,
    constants: IntegerConstantLookup,
) -> list[VbaToken] | None:
    if start >= end_exclusive:
        return None
    folded = fold_integer_expression_tokens(source, span, toks, start, end_exclusive, constants)
    if folded == 0:
        return toks[start:end_exclusive]
    # A comparison known False is 0: `5 / (n = 2)` with n = 1 (XLIDE issue #491).
    inner = toks[start:end_exclusive]
    depth = 0
    compares = False
    for tok in inner:
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        if depth == 0 and tok.raw_text in _COMPARISON_OPERATORS:
            compares = True
            break
    if compares and condition_value(inner, ConditionFacts(value=lambda lower: constants.get(lower))) is False:
        return inner
    if toks[start].raw_text == "(":
        close = match_paren_from(toks, start)
        if close == end_exclusive - 1:
            return _zero_divisor_expression(source, span, toks, start + 1, close, constants)
    if (
        end_exclusive == start + 2
        and toks[start].kind is TokenKind.OPERATOR
        and toks[start].raw_text in ("+", "-")
        and _is_zero_divisor_atom(toks[start + 1], constants)
    ):
        return [toks[start], toks[start + 1]]
    if end_exclusive == start + 1 and _is_zero_divisor_atom(toks[start], constants):
        return [toks[start]]
    if _zero_conversion_call_end(toks, start, constants) == end_exclusive - 1:
        return toks[start:end_exclusive]
    return None


def _zero_divisor_atom_token_group(
    toks: list[VbaToken], start: int, constants: IntegerConstantLookup
) -> list[VbaToken] | None:
    if start >= len(toks):
        return None
    first = toks[start]
    first_name = token_name(first)
    # Qualified conversions must precede the member-access branch, which
    # rejects a following parenthesis (XLIDE issue #898).
    close = _zero_conversion_call_end(toks, start, constants)
    if close is not None and _is_divisor_atom_boundary(_at(toks, close + 1)):
        return toks[start : close + 1]
    member = toks[start + 2] if start + 2 < len(toks) else None
    member_name = token_name(member) if member is not None else None
    if (
        first_name
        and start + 1 < len(toks)
        and toks[start + 1].raw_text == "."
        and member is not None
        and member_name
    ):
        # Only treat `first.member` as the complete divisor when nothing extends
        # the member-access chain past it; otherwise `a.Zero.Foo` / `a.Zero(i)`
        # would mis-match on the inner `a.Zero == 0` lookup.
        if not _is_divisor_atom_boundary(toks[start + 3] if start + 3 < len(toks) else None):
            return None
        return (
            [first, toks[start + 1], member]
            if constants.get(f"{first_name}.{member_name}".lower()) == 0
            else None
        )
    # A bare atom only stands alone when it is not itself a member-access head or
    # a call target (a following '.' or '(' means more of the expression follows).
    if _is_zero_divisor_atom(first, constants) and _is_divisor_atom_boundary(
        toks[start + 1] if start + 1 < len(toks) else None
    ):
        return [first]
    # `F()`, a Function of the module that returns 0 (XLIDE issue #448).
    if (
        first_name
        and _raw_at(toks, start + 1) == "("
        and _raw_at(toks, start + 2) == ")"
        and _is_divisor_atom_boundary(_at(toks, start + 3))
        and _raw_at(toks, start - 1) != "."
        and constants.get(f"{first_name}()") == 0
    ):
        return toks[start : start + 3]
    # `Int(0.9)`, `Fix(-0.9)`, `Round(0.5)`: a number made whole, 0 (XLIDE issue
    # #286). `Sign1(-1)`: a Function of the module that returns 0 for these
    # arguments (XLIDE issue #562).
    if first_name and _raw_at(toks, start + 1) == "(" and _raw_at(toks, start - 1) != ".":
        end = match_paren_from(toks, start + 1)
        call = toks[start : end + 1] if end > start else []
        if (
            len(call) > 0
            and _is_divisor_atom_boundary(_at(toks, end + 1))
            and evaluate_integer_constant_expression(" ".join(tok.raw_text for tok in call), constants) == 0
        ):
            return call
    return None


# Conversions to a whole-number type, which round their argument half to even.
_WHOLE_CONVERSIONS = frozenset({"cbyte", "cint", "clng", "clnglng", "clngptr"})
_FRACTIONAL_CONVERSIONS = frozenset({"csng", "cdbl", "ccur", "cdec"})


def _zero_conversion_call_end(
    toks: Sequence[VbaToken], start: int, constants: IntegerConstantLookup | None = None
) -> int | None:
    """Where a conversion of a literal that comes out 0 ends: `CLng(0)`,
    `CDbl(0)`, `CLng(0.4)`, which rounds to 0 (XLIDE issue #219, measured in Excel
    16.0; `10 / CDbl(0.4)` runs). A `VBA.` qualifier is allowed."""
    index = start
    if token_text(_at(toks, index)) == "vba" and _raw_at(toks, index + 1) == ".":
        index += 2
    name = token_text(_at(toks, index))
    whole = name in _WHOLE_CONVERSIONS
    if (
        (not whole and name not in _FRACTIONAL_CONVERSIONS)
        or _raw_at(toks, index + 1) != "("
        or _raw_at(toks, start - 1) == "."
    ):
        return None
    close = match_paren_from(toks, index + 1)
    inner = [] if close < 0 else toks[index + 2 : close]
    signed = len(inner) == 2 and inner[0].raw_text in ("-", "+")
    literal = inner[0] if len(inner) == 1 else inner[1] if signed else None
    # `CByte(o)` with o a Boolean still False, `CLng(v)` with v Empty (XLIDE issue #491).
    if literal is not None and constants is not None and _is_zero_divisor_atom(literal, constants):
        return close
    if literal is None or literal.kind not in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        return None
    if _is_zero_numeric_literal(literal):
        return close
    value = abs(js_number(_D_EXPONENT.sub("E", _NUMERIC_TYPE_SUFFIX_RE.sub("", literal.raw_text))))
    return close if whole and math.isfinite(value) and value <= 0.5 else None


def _is_divisor_atom_boundary(tok: VbaToken | None) -> bool:
    """True when tok terminates a divisor atom: end-of-tokens, or anything that is
    NOT a member-access dot or call-opening paren (either of those means the atom
    continues, so the group is not the whole divisor)."""
    if tok is None:
        return True
    return tok.raw_text != "." and tok.raw_text != "("


def _is_zero_divisor_atom(tok: VbaToken | None, constants: IntegerConstantLookup) -> bool:
    # False is 0 as a number: `1 \ False` raises 11 (XLIDE issue #458, measured in
    # Excel 16.0).
    if _is_zero_numeric_literal(tok) or (
        tok is not None and tok.kind is TokenKind.KEYWORD and token_text(tok) in ("false", "empty")
    ):
        return True
    name = token_name(tok) if tok is not None else None
    return name is not None and constants.get(name.lower()) == 0


def _is_zero_numeric_literal(tok: VbaToken | None) -> bool:
    if tok is None or tok.kind not in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
        return False
    normalized = _D_EXPONENT.sub("E", _TYPE_SUFFIX.sub("", tok.raw_text))
    hex_match = _HEX.match(normalized)
    if hex_match:
        return int(hex_match.group(1), 16) == 0
    octal_match = _OCTAL.match(normalized)
    if octal_match:
        return int(octal_match.group(1), 8) == 0
    if _FLOAT.match(normalized) is None:
        return False
    return float(normalized) == 0


def _absolute_token_group_span(base: Span, toks: list[VbaToken]) -> Span:
    return Span(base.start + toks[0].start, base.start + toks[-1].end)


# -- call-statement parenthesis rules --------------------------------------


def check_call_parens(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """A `Call` statement needs parentheses; a bare zero-arg call cannot use `()`.

    The standalone member-call form (`obj.Method()`) is reported too; a leading-dot
    member call (`.Method()` inside With) only fires when the member resolves against
    the receiver surface (the no-FP gate)."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            invalid_target = _invalid_explicit_call_target(source, stmt.span, module_signatures, source_names)
            if invalid_target is not None:
                name, span = invalid_target
                push(
                    "invalidExplicitCallTarget",
                    f"'{name}' cannot be used as the target of an explicit Call statement.",
                    span,
                )
                return
            at = explicit_call_statement_argument_without_parens(source, stmt.span)
            if at is not None:
                push(
                    "callRequiresParens",
                    "A Call statement requires parentheses around its argument list.",
                    at,
                )
            bare = _implicit_parenthesized_bare_callable_call(
                source, stmt.span, module_signatures, source_names
            )
            if bare is not None:
                name, span = bare
                push(
                    "callStatementForbidsParens",
                    _bare_call_forbids_parens_message(name, module_signatures, source_names),
                    span,
                )
            multi_arg = _implicit_parenthesized_multi_arg_call(
                source, stmt.span, module_signatures, source_names
            )
            if multi_arg is not None:
                name, span = multi_arg
                push(
                    "callStatementMultiArgParens",
                    f"A standalone call cannot enclose multiple arguments in parentheses; "
                    f"use 'Call {name}(...)' or remove the parentheses ('{name} arg1, arg2'). "
                    f"VBA rejects this form as a compile error.",
                    span,
                )
            implicit = _implicit_parenthesized_member_call(source, stmt.span, member_ctx)
            if implicit is not None:
                _name, span = implicit
                push(
                    "callStatementForbidsParens",
                    "Standalone zero-argument member calls cannot use empty parentheses unless "
                    "they are prefixed with Call or used in an expression.",
                    span,
                )

        return visitor

    return factory


def _implicit_parenthesized_member_call(
    source: str, span: Span, member_ctx: MemberCompletionContext
) -> tuple[str, Span] | None:
    """Port of implicitParenthesizedMemberCall: a standalone `obj.Method()` with empty
    parentheses. A leading-dot form (`.Method()` inside With) only counts when the
    member resolves against the receiver surface, the no-false-positive gate for an
    unknown With receiver."""
    call = standalone_empty_parenthesized_call_statement(source, span)
    if call is None or not call.is_member:
        return None
    if (
        call.starts_with_leading_dot
        and resolve_exact_member_completion(source, call.name, call.callee_end_offset, member_ctx)
        is None
    ):
        return None
    return (call.name, call.span)


def _bare_call_forbids_parens_message(
    name: str,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None,
) -> str:
    runtime = (
        resolve_runtime_function(name)
        if name.lower() not in module_signatures
        and not runtime_callable_source_shadowed(name, source_names)
        else None
    )
    if runtime is not None and not runtime_allows_explicit_call(runtime):
        return (
            f"Standalone '{runtime.name}()' cannot use empty parentheses in statement context; "
            f"use '{runtime.name}' as a statement or use it in an expression."
        )
    return (
        "Standalone zero-argument procedure calls cannot use empty parentheses unless they are "
        "prefixed with Call or used in an expression."
    )


def _invalid_explicit_call_target(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None,
) -> tuple[str, Span] | None:
    target = explicit_call_statement_target(source, span)
    if target is None:
        return None
    if target.name.lower() in module_signatures or runtime_callable_source_shadowed(
        target.name, source_names
    ):
        return None
    runtime = resolve_runtime_function(target.name)
    if runtime is None or runtime_allows_explicit_call(runtime):
        return None
    return (runtime.name, target.span)


def _implicit_parenthesized_bare_callable_call(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None,
) -> tuple[str, Span] | None:
    call = standalone_empty_parenthesized_call_statement(source, span)
    if call is None or call.is_member:
        return None
    signature = callable_signature_for(call.name, module_signatures, source_names)
    if signature is None or not callable_accepts_zero_arguments(signature):
        return None
    return (call.name, call.span)


def _implicit_parenthesized_multi_arg_call(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope | None,
) -> tuple[str, Span] | None:
    """Port of implicitParenthesizedMultiArgCall: a standalone `mySub2("a", "b", "c")`
    wraps a multi-argument list in parentheses without `Call` (the VBE "Expected: ="
    compile error). Scoped to a callee that resolves to a known procedure so unknown
    names (which could be array indexing or external references) stay silent:

    * bare names bind to same-module/unique-exported project Sub/Function/Declare;
    * `Module.Proc(...)` binds to an exported standard-module procedure through its
      deterministic qualified key (the same resolution the argument-count rule uses).

    A member call is reported whatever its receiver: `(1, 2)` is no expression, so
    the statement is a Syntax error for `c.Calc (1, 2)` on a class, `.Calc (1, 2)`
    in a With, and `ActiveSheet.Range ("A1", "B2")` (XLIDE issue #224, measured).
    The single-argument ByVal-grouping form is excluded by the >= 2 argument guard
    in the shared helper."""
    call = standalone_multi_arg_parenthesized_call_statement(source, span)
    if call is None:
        return None
    if call.is_member:
        return (f"{call.qualifier}.{call.name}" if call.qualifier else call.name, call.span)
    if callable_signature_for(call.name, module_signatures, source_names) is None:
        return None
    return (call.name, call.span)


# -- expression-call parenthesis rule --------------------------------------


def check_expression_call_parens(
    source: str,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """A Function used inside an expression must parenthesize its argument list."""
    bare, qualified = _expression_callable_function_names(symbols, project_procedures)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            hit = _parenless_expression_call(source, stmt.span, bare, qualified, source_names)
            if hit is not None:
                name, span = hit
                push(
                    "expressionCallRequiresParens",
                    f"Function call arguments in an expression must be enclosed in "
                    f"parentheses: use '{name}(...)'.",
                    span,
                )

        return visitor

    return factory


def _expression_callable_function_names(
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
) -> tuple[set[str], set[str]]:
    bare: set[str] = set()
    qualified: set[str] = set()
    for member in symbols.root.children or []:
        if member.kind in (VbaSymbolKind.FUNCTION, VbaSymbolKind.PROPERTY_GET):
            bare.add(member.name.lower())
    for key, candidates in (project_procedures or {}).items():
        if len(candidates) != 1 or candidates[0].kind != "function":
            continue
        if "." in key:
            qualified.add(key)
        elif key not in bare:
            bare.add(key)
    return bare, qualified


def _parenless_expression_call(
    source: str,
    span: Span,
    bare: set[str],
    qualified: set[str],
    source_names: SourceNameScope | None,
) -> tuple[str, Span] | None:
    toks = statement_tokens(source, span)
    if not toks or _is_non_assignment_statement_leader(_statement_head_word(toks)):
        return None
    eq = top_level_operator_index(toks, "=")
    if eq < 0:
        return None
    for i in range(eq + 1, len(toks) - 1):
        tok = toks[i]
        name = token_name(tok)
        if not name or not _is_expression_callable_at(toks, i, name, bare, qualified, source_names):
            continue
        if i > eq + 1 and toks[i - 1].raw_text == ".":
            qualifier = token_name(toks[i - 2]) if i >= 2 else None
            if not qualifier or qualified_procedure_key(qualifier, name) not in qualified:
                continue  # object member calls need receiver typing
        nxt = toks[i + 1]
        if not _is_parenless_argument_start(nxt):
            continue
        gap = source[span.start + tok.end : span.start + nxt.start]
        if not any(c.isspace() for c in gap):
            continue
        return (name, Span(span.start + tok.start, span.start + tok.end))
    return None


def _is_expression_callable_at(
    toks: list[VbaToken],
    index: int,
    name: str,
    bare: set[str],
    qualified: set[str],
    source_names: SourceNameScope | None,
) -> bool:
    if index > 1 and toks[index - 1].raw_text == ".":
        qualifier = token_name(toks[index - 2])
        return qualifier is not None and qualified_procedure_key(qualifier, name) in qualified
    if index > 0 and toks[index - 1].raw_text == ".":
        return False
    if bare_callable_source_shadowed(name, source_names):
        return False
    if name.lower() in bare:
        return True
    if runtime_callable_source_shadowed(name, source_names):
        return False
    runtime = resolve_runtime_function(name)
    return runtime is not None and runtime.kind == "function"


_INFIX_KEYWORDS = frozenset({"and", "or", "xor", "eqv", "imp", "is", "mod"})
_NON_ASSIGNMENT_LEADERS = frozenset(
    {"if", "elseif", "for", "do", "loop", "while", "select", "case"}
)


def _is_parenless_argument_start(tok: VbaToken | None) -> bool:
    if tok is None:
        return False
    if tok.kind in (
        TokenKind.IDENTIFIER,
        TokenKind.BRACKETED_IDENTIFIER,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    ):
        return True
    if tok.kind is TokenKind.KEYWORD:
        return tok.raw_text.lower() not in _INFIX_KEYWORDS
    return False


def _statement_head_word(toks: Sequence[VbaToken]) -> str:
    """The word the statement starts with, past any line number or label."""
    first = first_executable_token_index(toks)
    return token_text(toks[first] if first < len(toks) else None)


def _is_non_assignment_statement_leader(word: str) -> bool:
    return word in _NON_ASSIGNMENT_LEADERS


# -- invalid expression syntax ---------------------------------------------

_NON_UNARY_BINARY_OPERATORS = frozenset(
    {
        "*", "/", "\\", "^", "&", "=", "<", ">", "<=", ">=", "<>", ":=",
        "like", "is", "and", "or", "xor", "eqv", "imp", "mod",
    }
)


def check_invalid_expression_syntax(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Incomplete member access, the unsupported `?` operator, and invalid operator runs."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_scalar_type(name: str) -> SourceDeclaredType:
            return declared_type_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.MEMBER_RECEIVER
            )

        def visitor(stmt: LeafStatementNode) -> None:
            incomplete = incomplete_member_access(
                source, stmt.span, scalar_types=env, resolve_scalar_type=resolve_scalar_type
            )
            if incomplete is not None:
                push(
                    "invalidExpressionSyntax",
                    "Incomplete member access: type a member name after '.'.",
                    incomplete,
                )
                return
            unsupported = _unsupported_question_mark_operator(source, stmt.span)
            if unsupported is not None:
                push(
                    "invalidExpressionSyntax",
                    "VBA does not support the '?' conditional operator in code modules; use "
                    "If...Then...Else, or IIf(...) only when both branches are safe to evaluate.",
                    unsupported,
                )
                return
            hit = _invalid_operator_sequence(source, stmt.span)
            if hit is not None:
                text, hit_span, hit_message = hit
                push(
                    "invalidExpressionSyntax",
                    hit_message
                    if hit_message is not None
                    else f"Invalid operator sequence '{text}'; this will fail to compile as a syntax error.",
                    hit_span,
                )
                return
            juxtaposed = _juxtaposed_rhs_values(source, stmt.span)
            if juxtaposed is not None:
                text, hit_span = juxtaposed
                push(
                    "invalidExpressionSyntax",
                    f"Unexpected '{text}' after a complete expression; expected end of "
                    f"statement. This will fail to compile as a syntax error.",
                    hit_span,
                )

        return visitor

    return factory


def incomplete_member_access(
    source: str,
    span: Span,
    *,
    include_leading_dot: bool = False,
    scalar_types: Mapping[str, str] | None = None,
    resolve_scalar_type: Callable[[str], SourceDeclaredType] | None = None,
) -> Span | None:
    toks = statement_tokens(source, span)
    for i, tok in enumerate(toks):
        if tok.raw_text != ".":
            continue
        if i == 0 and not include_leading_dot:
            continue
        nxt = toks[i + 1] if i + 1 < len(toks) else None
        if nxt is not None and token_name(nxt):
            continue
        receiver_name = token_name(toks[i - 1]) if i > 0 else None
        if receiver_name:
            resolved = resolve_scalar_type(receiver_name) if resolve_scalar_type else None
            as_type = (
                resolved.as_type
                if resolved is not None and resolved.resolved
                else (scalar_types.get(receiver_name.lower()) if scalar_types else None)
            )
            normalized = normalize_type(as_type)
            if normalized and is_known_scalar_type(normalized):
                continue
        return absolute_span(span, tok)
    return None


def _unsupported_question_mark_operator(source: str, span: Span) -> Span | None:
    for tok in statement_tokens(source, span):
        if tok.kind is TokenKind.OPERATOR and tok.raw_text == "?":
            return absolute_span(span, tok)
    return None


def is_non_unary_binary_operator(tok: VbaToken | None) -> bool:
    if tok is None:
        return False
    # VBA word operators (And/Or/Xor/Eqv/Imp/Mod/Like/Is) lex as keyword tokens,
    # so accept those alongside symbolic operator tokens (e.g. ':=') by matching
    # on the lowercased text rather than the token kind.
    if tok.kind is not TokenKind.OPERATOR and tok.kind is not TokenKind.KEYWORD:
        return False
    return token_text(tok) in _NON_UNARY_BINARY_OPERATORS


_JUXTAPOSABLE_VALUE_KINDS = frozenset(
    {
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.DATE_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.IDENTIFIER,
        TokenKind.BRACKETED_IDENTIFIER,
    }
)


def _is_juxtaposable_value_start(tok: VbaToken | None) -> bool:
    return tok is not None and tok.kind in _JUXTAPOSABLE_VALUE_KINDS


def _ends_juxtaposable_value(tok: VbaToken | None) -> bool:
    if tok is None:
        return False
    # A digit run glued to `&` lexes as a &-suffixed integer literal, but the
    # VBE can read that `&` as CONCATENATION - it does when the digits overflow
    # Long (VBE oracle suffix_long_amp_glued_concat_accepted: `s = 3000000000&"x"`
    # is accepted) - so a &-suffixed integer literal never provably ends a
    # value. The in-range form (`n = 5& 1`) is under-reported by design: a
    # missed diagnostic beats a false positive on the oracle-verified concat.
    if tok.kind is TokenKind.INTEGER_LITERAL and tok.raw_text.endswith("&"):
        return False
    return tok.kind in _JUXTAPOSABLE_VALUE_KINDS or tok.raw_text in (")", "]")


def _juxtaposed_rhs_values(source: str, span: Span) -> tuple[str, Span] | None:
    """Detects two juxtaposed value expressions in an assignment RHS - a complete
    value (literal / identifier / call / index) immediately followed by another
    value starter with no operator between, e.g. `n = 1 n 1` or
    `n = 1 MsgBox("hello") 1`. That is a VBE "Expected: end of statement" syntax
    error which the lenient parser otherwise silently drops to a raw statement.

    Scoped to the TOP LEVEL of an assignment RHS (a top-level standalone `=` on a
    statement that is not a non-assignment leader) so it cannot misfire on:
    implicit call statements (`MsgBox x` - no `=`), a call written with a space
    (`Foo (x)` - the next token is `(`, not a value start), jagged-array access
    (`arr(1)(2)` - `(` again), a trailing type-suffix/operator (`Count&` - `&` is
    not a value start), or anything inside parentheses (depth > 0 is skipped)."""
    toks = statement_tokens(source, span)
    if len(toks) == 0 or _is_non_assignment_statement_leader(_statement_head_word(toks)):
        return None
    eq = top_level_operator_index(toks, "=")
    if eq < 0:
        return None
    at = juxtaposed_value_index(toks, eq + 1)
    return None if at < 0 else (toks[at].raw_text, absolute_span(span, toks[at]))


def juxtaposed_value_index(toks: Sequence[VbaToken], from_: int) -> int:
    """The index of a value that follows a complete value with no operator
    between, `asdf qwer` or `1 n`, from `from_` on at the top level; -1 when there
    is none. A Const's value is read the same way (XLIDE issue #234)."""
    depth = 0
    for i in range(from_, len(toks) - 1):
        raw = toks[i].raw_text
        if raw in ("(", "["):
            depth += 1
            continue
        if raw in (")", "]"):
            depth = depth - 1 if depth > 0 else 0
        if depth != 0:
            continue
        if _ends_juxtaposable_value(toks[i]) and _is_juxtaposable_value_start(toks[i + 1]):
            return i + 1
    return -1


def is_glued_type_suffix_ampersand(toks: Sequence[VbaToken], index: int) -> bool:
    """Whether the `&` at `index` is glued to a name before it, which makes it the
    name's Long type-declaration character (`total&`) rather than the concatenation
    operator. `s$`, `n%`, `x!`, `d#` and `c@` lex the same way; only `&` doubles as
    an operator, so only it needs asking."""
    if not 0 <= index < len(toks) or index < 1:
        return False
    tok = toks[index]
    prev = toks[index - 1]
    return (
        tok.kind is TokenKind.OPERATOR
        and tok.raw_text == "&"
        and prev.end == tok.start
        and token_name(prev) is not None
    )


def _invalid_operator_sequence(source: str, span: Span) -> tuple[str, Span, str | None] | None:
    """(text, span, message) of the first impossible operator run; the message is
    None where the caller words it from the text."""
    toks = statement_tokens(source, span)
    # A Case statement's Is-comparison clause (MS-VBAL 5.4.2.10, `Case Is > 5`)
    # uses `Is` as grammar, not as the object-identity operator, so the
    # word-operator scan would mis-read `Is >` as an impossible operator run.
    # Case clauses are grammar, not value expressions; skip the whole statement
    # (the Select/Case rules own its structure).
    head = first_executable_token_index(toks)
    if head < len(toks) and token_text(toks[head]) == "case":
        return None
    i = 0
    while i < len(toks):
        tok = toks[i]
        # `Not` with nothing after it to negate: a line holding only `Not`, or
        # `x = Not` (XLIDE issue #234, measured in Excel 16.0: Syntax error).
        if (
            token_text(tok) == "not"
            and tok.kind is TokenKind.KEYWORD
            and (i == len(toks) - 1 or toks[i + 1].kind is TokenKind.COMMENT)
        ):
            return (
                tok.raw_text,
                Span(span.start + tok.start, span.start + tok.end),
                "'Not' has nothing after it to negate. This is a VBE compile error: Syntax error.",
            )
        if not is_non_unary_binary_operator(tok):
            i += 1
            continue
        # `total& = 3`: an `&` glued to the name before it is the Long
        # type-declaration character, not concatenation. The VBE reads it that way
        # whatever follows - `a& b` is a syntax error there, `a &b` is a
        # concatenation (XLIDE issue #100, measured in Excel 16.0).
        if is_glued_type_suffix_ampersand(toks, i):
            i += 1
            continue
        # `x^=5` assigns the LongLong x^ (XLIDE issue #369, measured in Excel 16.0);
        # elsewhere a glued `^` is the power operator, `a^2`.
        if (
            tok.raw_text == "^"
            and i == first_executable_token_index(toks) + 1
            and toks[i - 1].end == tok.start
            and token_name(toks[i - 1]) is not None
            and _raw_at(toks, i + 1) == "="
        ):
            i += 1
            continue
        # `a < > b` is one relational operator written as two tokens (MS-VBAL
        # 5.6.9.5), and the VBE reads it as `a <> b` (XLIDE issue #87).
        relational = relational_operator_at(toks, i)
        operator_end = i + (relational[1] if relational is not None else 1) - 1
        end = operator_end
        while end + 1 < len(toks) and is_non_unary_binary_operator(toks[end + 1]):
            end += 1
        if end > operator_end or operator_end == len(toks) - 1:
            first = toks[i]
            last = toks[end]
            return (
                source[span.start + first.start : span.start + last.end],
                Span(span.start + first.start, span.start + last.end),
                None,
            )
        i = operator_end + 1
    return None
