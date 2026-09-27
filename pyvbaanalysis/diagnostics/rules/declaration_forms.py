"""Rule family: declaration forms the VBE refuses while compiling (XLIDE issue #124).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/declarationForms.ts. Each
was measured in Excel 16.0 (build 20326, 2026-09-25/26) with the message quoted.

- array-parameter-form: `ByVal a() As Long` -> "Array argument must be ByRef";
  `Optional a() As Long` -> "Optional argument must be Variant or intrinsic type
  with a default value".
- parameter-default-type-mismatch: `Optional ByVal i As Integer = 40000`,
  `Optional ByVal b As Byte = 256` -> "Overflow".
- const-overflow: an Enum member `eA = 3000000000#` -> "Overflow" (Enum members
  are Long).
- duplicate-deftype: `DefLng A-Z` followed by `DefStr S` -> "Duplicate Deftype
  statement": a letter may be given a default type once.
- bracketed-variable-name: `Dim [my var] As Long` -> "Syntax error". Brackets make
  a foreign name for an Enum member (`[Two Words] = 2` compiles) but not for a
  variable.
"""

from __future__ import annotations

import math
from collections.abc import Sequence

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import parse_vba_integer_literal
from ...js_compat import js_number, js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    EnumNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
)
from ...types.type_names import normalize_type, numeric_literal_bounds
from ..context import PushFn
from ..walker import (
    absolute_span,
    active_module_members,
    declared_name_span,
    for_each_variable_group,
    is_inactive_node,
    raw_expression_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .shared import DEFTYPE_KEYWORDS

_LONG_RANGE_MIN = -2147483648
_LONG_RANGE_MAX = 2147483647

# A type-declaration suffix a float literal may end with (Single, Double, Currency).
_FLOAT_TYPE_SUFFIXES = ("!", "#", "@")


def check_declaration_forms(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    claimed_letters: dict[str, str] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for param in member.params:
                _check_parameter(source, param, push)
            for_each_variable_group(
                member.body, lambda group: _check_group_names(source, group, push), activity
            )
        elif isinstance(member, VariableGroupNode):
            _check_group_names(source, member, push)
        elif isinstance(member, EnumNode):
            for item in member.members:
                if item.value_raw is None or is_inactive_node(activity, item):
                    continue
                value = _literal_number(raw_expression_tokens(item.value_raw))
                if value is not None and (value < _LONG_RANGE_MIN or value > _LONG_RANGE_MAX):
                    push(
                        "constOverflow",
                        f"Enum member '{item.name}' is {js_trim(item.value_raw)}, outside the Long "
                        f"range {_LONG_RANGE_MIN} to {_LONG_RANGE_MAX} an Enum member holds. "
                        "This is a VBE compile error: Overflow.",
                        declared_name_span(source, item.span, item.name),
                    )
        elif isinstance(member, StatementNode):
            _check_deftype(source, member.span, claimed_letters, push)


def _check_group_names(source: str, group: VariableGroupNode, push: PushFn) -> None:
    for decl in group.declarations:
        _check_bracketed_name(source, decl.name, decl.name_span, push)


def _check_parameter(source: str, param: ParameterNode, push: PushFn) -> None:
    name_span = param.name_span if param.name_span is not None else param.span
    if param.is_array and param.by_val and not param.param_array:
        push(
            "arrayParameterForm",
            f"Array parameter '{param.name}' cannot be ByVal: an array argument must be ByRef.",
            name_span,
        )
    if param.is_array and param.optional:
        push(
            "arrayParameterForm",
            f"Optional parameter '{param.name}' cannot be an array: an Optional argument must be "
            "Variant or an intrinsic type with a default value.",
            name_span,
        )
    if not param.optional or param.default_raw is None or param.is_array:
        return
    type_ = normalize_type(param.as_type)
    bounds = numeric_literal_bounds(type_) if type_ else None
    if bounds is None:
        return
    value = _literal_number(raw_expression_tokens(param.default_raw))
    if value is None:
        return
    stored = value if type_ in ("single", "double", "currency") else _js_math_round(value)
    if stored < bounds.min or stored > bounds.max:
        default = js_trim(param.default_raw)
        at = source.find(default, param.span.start)
        span = Span(at, at + len(default)) if 0 <= at < param.span.end else param.span
        push(
            "parameterDefaultTypeMismatch",
            f"Optional parameter '{param.name}' is declared As {param.as_type}, whose range is "
            f"{bounds.min} to {bounds.max}; its default {default} does not fit. "
            "This is a VBE compile error: Overflow.",
            span,
        )


def _literal_number(toks: Sequence[VbaToken]) -> float | None:
    """The number a plain signed literal denotes, or None for anything else."""
    sign = 1
    rest = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    if rest and (rest[0].raw_text == "-" or rest[0].raw_text == "+"):
        sign = -1 if rest[0].raw_text == "-" else 1
        rest = rest[1:]
    if len(rest) != 1:
        return None
    literal = rest[0]
    if literal.kind is TokenKind.INTEGER_LITERAL:
        parsed = parse_vba_integer_literal(literal.raw_text)
        return None if parsed is None else sign * parsed
    if literal.kind is TokenKind.FLOAT_LITERAL:
        text = literal.raw_text.replace("d", "E").replace("D", "E")
        if text[-1:] in _FLOAT_TYPE_SUFFIXES:
            text = text[:-1]
        number = js_number(text)
        return sign * number if math.isfinite(number) else None
    return None


def _check_bracketed_name(source: str, name: str, name_span: Span | None, push: PushFn) -> None:
    if name_span is None or not 0 <= name_span.start < len(source):
        return
    if source[name_span.start] == "[":
        push(
            "bracketedVariableName",
            f"'[{name}]' is not a variable name: brackets make a foreign name only for an Enum "
            "member or a member of another object. This is a VBE compile error: Syntax error.",
            name_span,
        )


def _check_deftype(source: str, span: Span, claimed: dict[str, str], push: PushFn) -> None:
    """`DefLng A-Z` then `DefStr S`: every letter a Deftype names must be new.

    Letters are tracked across the module's Deftype statements in order.
    """
    toks = statement_tokens_after_leading_label(source, span)
    head = token_text(toks[0] if toks else None)
    if head not in DEFTYPE_KEYWORDS:
        return
    first = toks[0]
    statement = first.canonical_text if first.canonical_text is not None else first.raw_text
    letters: list[str] = []
    for i in range(1, len(toks)):
        name = token_name(toks[i])
        if not name or len(name) != 1:
            continue
        upper = name.upper()
        if toks[i - 1].raw_text == "-" and letters:
            start = ord(letters[-1][0])
            for code in range(start + 1, ord(upper[0]) + 1):
                letters.append(chr(code))
            continue
        letters.append(upper)
    for letter in letters:
        earlier = claimed.get(letter)
        if earlier:
            push(
                "duplicateDeftype",
                f"Letter '{letter}' already has a default type from '{earlier}' above; a letter "
                "takes one Deftype statement. This is a VBE compile error: Duplicate Deftype "
                "statement.",
                absolute_span(span, first),
            )
            return
    for letter in letters:
        claimed[letter] = statement


def _js_math_round(value: float) -> float:
    """JavaScript's Math.round: the nearest integer, a tie going toward +Infinity.

    Python's round() breaks a tie to the even neighbor (round(2.5) == 2).
    """
    floor = math.floor(value)
    return floor + 1 if value - floor >= 0.5 else floor
