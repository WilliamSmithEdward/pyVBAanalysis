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
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...completion.member_access import (
    MemberCompletionContext,
    resolve_exact_member_completion,
)
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    resolve_raw_integer_constants,
)
from ...host.host_model import HostObjectModel
from ...lexer.token_helpers import match_paren_from, relational_operator_at
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import (
    BodyNode,
    IfBlockNode,
    IfBranchKind,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
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
    procedure_symbol_for,
    type_environment_for,
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
from ..const_expr import (
    collect_module_literal_integer_constants,
    fold_integer_expression_tokens,
)
from ..context import PushFn, statement_tokens
from ..known_locals import ConstantOrKnownLocalLookup, known_local_literal_values
from ..walker import (
    ProcedureStatementVisitor,
    absolute_span,
    bare_assignment_target,
    first_executable_token_index,
    raw_expression_tokens,
    token_name,
    token_text,
    top_level_operator_index,
)


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

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        constants = procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        # A local the procedure never assigns is 0, and one whose every assignment
        # is `d = 0` is 0 too (XLIDE issue #119): `10 / d` raises 11.
        lookup = ConstantOrKnownLocalLookup(
            constants, known_local_literal_values(source, member, symbols, activity)
        )
        guards = _division_guard_ranges(member.body, activity)

        def visitor(stmt: LeafStatementNode) -> None:
            for message, span in _division_by_zero_divisors(source, stmt.span, lookup, guards):
                push("divisionByZero", message, span)

        return visitor

    return factory


# The operators that read both sides as numbers when either is a number.
_ARITHMETIC_OR_COMPARISON = frozenset({"+", "-", "*", "/", "\\", "^", "=", "<", ">", "<=", ">=", "<>"})
_ALWAYS_COERCING = frozenset({"-", "*", "/", "\\", "^"})
_OPERAND_END_KINDS = frozenset(
    {
        TokenKind.IDENTIFIER,
        TokenKind.KEYWORD,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    }
)


def check_string_arithmetic_operands(
    source: str,
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

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        known = known_local_literal_values(source, member, symbols, activity)

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
            return None

        def numeric(tok: VbaToken | None) -> bool:
            if tok is None:
                return False
            if tok.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
                return True
            name = token_name(tok)
            if not name:
                return False
            local = known.get(name.lower())
            if local is not None and local.kind == "number":
                return True
            type_ = normalize_type(env.get(name.lower()))
            return type_ is not None and is_numeric_type(type_)

        def visitor(stmt: LeafStatementNode) -> None:
            toks = statement_tokens(source, stmt.span)
            if _statement_head_word(toks) == "const":
                return
            # An assignment to a numeric variable is the assignment rule's: it already
            # names the target, and one report per line is enough.
            bare = bare_assignment_target(source, stmt.span)
            if bare is not None:
                target_type = normalize_type(env.get(bare[0].lower()))
                if target_type and is_numeric_type(target_type):
                    return
            for i, tok in enumerate(toks):
                word = token_text(tok)
                left = toks[i - 1] if i >= 1 else None
                right = toks[i + 1] if i + 1 < len(toks) else None

                def report(operand: VbaToken, what: str, operator: VbaToken = tok) -> None:
                    push(
                        "stringArithmeticCoercion",
                        f"Operator '{operator.raw_text}' coerces {what} to a number. This will raise "
                        "Run-time error '13': Type mismatch.",
                        absolute_span(stmt.span, operand),
                    )

                is_binary = (
                    tok.raw_text in _ARITHMETIC_OR_COMPARISON
                    if tok.kind is TokenKind.OPERATOR
                    else word == "mod"
                )
                left_ends_operand = left is not None and (
                    left.kind in _OPERAND_END_KINDS or left.raw_text == ")"
                )
                if word == "not" and not left_ends_operand:
                    what = nonnumeric_string(right)
                    if what and right is not None:
                        report(right, what)
                    continue
                if not is_binary:
                    continue
                if not left_ends_operand:
                    # Unary `-"abc"` (a leading `+` too).
                    if tok.raw_text in ("-", "+"):
                        what = nonnumeric_string(right)
                        if what and right is not None:
                            report(right, what)
                    continue
                always_coerces = tok.raw_text in _ALWAYS_COERCING or word == "mod"
                left_string = nonnumeric_string(left)
                right_string = nonnumeric_string(right)
                if always_coerces:
                    if left_string and left is not None:
                        report(left, left_string)
                    elif right_string and right is not None:
                        report(right, right_string)
                    continue
                # `+` and comparisons: a string against a NUMBER.
                if left_string and left is not None and numeric(right):
                    report(left, left_string)
                elif right_string and right is not None and numeric(left):
                    report(right, right_string)

        return visitor

    return factory


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
    source: str, span: Span, constants: IntegerConstantLookup, guards: Sequence[_DivisionGuard]
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
            divisor = _fractional_divisor_rounding_to_zero(toks, i + 1, operator)
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
    toks: Sequence[VbaToken], start: int, operator: str
) -> list[VbaToken] | None:
    """`\\` and `Mod` round their operands to whole numbers first, with banker's
    rounding, so a literal divisor below 0.5 - or exactly 0.5 - is zero to them:
    `5 \\ 0.4` and `5 Mod 0.5` raise 11 (XLIDE issue #119, measured in Excel 16.0)."""
    if operator == "/":
        return None
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
    return None


def _zero_divisor_atom_token_group(
    toks: list[VbaToken], start: int, constants: IntegerConstantLookup
) -> list[VbaToken] | None:
    if start >= len(toks):
        return None
    first = toks[start]
    first_name = token_name(first)
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
    return None


def _is_divisor_atom_boundary(tok: VbaToken | None) -> bool:
    """True when tok terminates a divisor atom: end-of-tokens, or anything that is
    NOT a member-access dot or call-opening paren (either of those means the atom
    continues, so the group is not the whole divisor)."""
    if tok is None:
        return True
    return tok.raw_text != "." and tok.raw_text != "("


def _is_zero_divisor_atom(tok: VbaToken | None, constants: IntegerConstantLookup) -> bool:
    if _is_zero_numeric_literal(tok):
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

    Object member/property calls (`obj.Method(a, b)`) are deliberately deferred: a
    non-empty single-argument member form is legal (`ActiveSheet.Range("A1")`) and
    multi-argument member/default-member forms are unproven. The single-argument
    ByVal-grouping form is excluded by the >= 2 argument guard in the shared helper."""
    call = standalone_multi_arg_parenthesized_call_statement(source, span)
    if call is None:
        return None
    if call.is_member:
        if not call.qualifier or qualified_procedure_key(call.qualifier, call.name) not in module_signatures:
            return None
        return (f"{call.qualifier}.{call.name}", call.span)
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
                text, hit_span = hit
                push(
                    "invalidExpressionSyntax",
                    f"Invalid operator sequence '{text}'; this will fail to compile as a syntax error.",
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


def _is_non_unary_binary_operator(tok: VbaToken | None) -> bool:
    # VBA word operators (And/Or/Xor/Eqv/Imp/Mod/Like/Is) lex as keyword tokens,
    # so accept those alongside symbolic operator tokens (e.g. ':=') by matching
    # on the lowercased text rather than the token kind.
    return (
        tok is not None
        and tok.kind in (TokenKind.OPERATOR, TokenKind.KEYWORD)
        and token_text(tok) in _NON_UNARY_BINARY_OPERATORS
    )


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
    depth = 0
    for i in range(eq + 1, len(toks) - 1):
        raw = toks[i].raw_text
        if raw in ("(", "["):
            depth += 1
            continue
        if raw in (")", "]"):
            depth = depth - 1 if depth > 0 else 0
        if depth != 0:
            continue
        nxt = toks[i + 1]
        if _ends_juxtaposable_value(toks[i]) and _is_juxtaposable_value_start(nxt):
            return (nxt.raw_text, absolute_span(span, nxt))
    return None


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


def _invalid_operator_sequence(source: str, span: Span) -> tuple[str, Span] | None:
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
        if not _is_non_unary_binary_operator(toks[i]):
            i += 1
            continue
        # `total& = 3`: an `&` glued to the name before it is the Long
        # type-declaration character, not concatenation. The VBE reads it that way
        # whatever follows - `a& b` is a syntax error there, `a &b` is a
        # concatenation (XLIDE issue #100, measured in Excel 16.0).
        if is_glued_type_suffix_ampersand(toks, i):
            i += 1
            continue
        # `a < > b` is one relational operator written as two tokens (MS-VBAL
        # 5.6.9.5), and the VBE reads it as `a <> b` (XLIDE issue #87).
        relational = relational_operator_at(toks, i)
        operator_end = i + (relational[1] if relational is not None else 1) - 1
        end = operator_end
        while end + 1 < len(toks) and _is_non_unary_binary_operator(toks[end + 1]):
            end += 1
        if end > operator_end or operator_end == len(toks) - 1:
            first = toks[i]
            last = toks[end]
            return (
                source[span.start + first.start : span.start + last.end],
                Span(span.start + first.start, span.start + last.end),
            )
        i = operator_end + 1
    return None
