"""Rule family: arrays one step removed (XLIDE issue #248).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/typeFieldArrays.ts.
Every case was measured in Excel 16.0 (build 20326, 2026-09-30).

 - array-subscript-out-of-bounds: a subscript outside an array field of a
   user-defined type, `t.arr(5)` on `arr(3) As Long`, at any depth,
   `t.kids(0).vals(9)`, and inside `With t`; UBound or LBound of a dimension
   the field lacks; and a dynamic array field nothing has ReDimmed, or that
   Erase emptied, `t.dyn(0)`, or past the bounds its last ReDim gave it. Each
   raises 9.
 - wrong-number-of-dimensions: `t.arr(1, 1)` on a field of one dimension, a
   compile error.
 - variable-required: `Len(a)` or `LenB(a)` of an array, a variable or a
   field, parentheses or not: "Variable required - can't assign to this
   expression" at compile time.
 - variant-value-misuse: `VBA.Len(a)` is the library function, which takes the
   array as a Variant: 13 for any array but a Byte array.

A dynamic field is followed only on a local of the Type, not Static, from its
Dim, as type_member_state.py follows it.
"""

from __future__ import annotations

import dataclasses
import math
import re
from collections.abc import Callable, Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import IntegerConstantLookup, resolve_raw_integer_constants
from ...host.host_model import HostObjectModel
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...runtime.vba_runtime import resolve_runtime_function
from ...symbols.symbol_model import ModuleSymbols, VbaSymbol, VbaSymbolKind
from ...types.type_inference import known_local_literal_values_at, type_environment_for, with_known_locals
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import infer_expression_type
from ..callable_signatures import (
    SourceNameScope,
    build_module_type_signatures,
    procedure_integer_constant_lookup,
    runtime_callable_source_shadowed,
    source_name_scope_for,
)
from ..const_expr import collect_module_literal_integer_constants
from ..context import PushFn
from ..type_fields import (
    FieldStep,
    ModuleTypes,
    WithSubject,
    field_chain,
    is_fixed_array_field,
    module_types,
    type_key,
    type_root_at,
    variable_symbol_in,
    walk_with_subjects,
)
from ..type_member_state import MemberState, MemberStatesAt, is_array_bounds, type_member_states_at
from ..walker import (
    active_module_members,
    match_paren_from,
    pluralize_count,
    raw_expression_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, SubscriptHit, module_option_base, shape_subscript_violation
from .shared import is_bare_or_vba_qualified_intrinsic_call

_SUBSCRIPT_ERROR = "This will raise Run-time error '9': Subscript out of range."

_ValueType = Callable[[Sequence[VbaToken]], "str | None"]


def _no_value_type(_value: Sequence[VbaToken]) -> str | None:
    return None


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def check_type_field_arrays(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_integer_constants: Mapping[str, str | None] | None = None,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
    host_model: HostObjectModel | None = None,
) -> None:
    types = module_types(source, mod, activity)
    module_signatures = build_module_type_signatures(symbols)
    option_base = module_option_base(mod, activity)
    module_constants = collect_module_literal_integer_constants(
        mod, activity, resolve_raw_integer_constants(project_integer_constants or {}, {})
    )
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        env = type_environment_for(symbols, member)

        def value_type(
            value: Sequence[VbaToken],
            env: Mapping[str, str] = env,
            source_names: SourceNameScope = source_names,
        ) -> str | None:
            inferred = infer_expression_type(list(value), 0, env, module_signatures, source_names, source=source)
            return normalize_type(inferred.type_ if inferred is not None else None)

        constants = procedure_integer_constant_lookup(
            member, module_constants, symbols, project_visible_symbols, activity, host_model
        )
        values_at = known_local_literal_values_at(source, member, symbols, activity)
        states_at: MemberStatesAt | None = (
            None if not types else type_member_states_at(source, symbols, member, types, activity, option_base)
        )

        def visit(
            stmt: LeafStatementNode,
            subject: WithSubject | None,
            member: ProcedureNode = member,
            source_names: SourceNameScope = source_names,
            value_type: _ValueType = value_type,
            constants: IntegerConstantLookup = constants,
            values_at: Callable[..., Mapping[str, object]] = values_at,
            states_at: MemberStatesAt | None = states_at,
        ) -> None:
            toks = statement_tokens_after_leading_label(source, stmt.span)
            lookup = with_known_locals(constants, values_at(stmt))  # type: ignore[arg-type]
            for span, message, rule in _len_of_arrays(
                stmt.span, toks, symbols, member, types, subject, source_names, value_type
            ):
                push(rule, message, span)
            if states_at is None or token_text(_at(toks, 0)) in ("redim", "erase"):
                return  # a ReDim's bounds are not subscripts
            for i in range(len(toks)):
                root = type_root_at(toks, i, symbols, member, types, subject)
                states = states_at(stmt, stmt.span.start + toks[i].start) if root is not None else None
                hit = (
                    _chain_violation(stmt.span, toks, i, field_chain(toks, root, types), states, lookup)
                    if root is not None and states is not None
                    else None
                )
                if hit is not None:
                    push(hit.rule or "arraySubscriptOutOfBounds", hit.message, hit.span)

        walk_with_subjects(source, member.body, activity, symbols, member, types, None, visit)


_NO_ELEMENTS = (
    "has no elements here: it is a dynamic array field that nothing has ReDimmed, or that Erase emptied. "
    f"{_SUBSCRIPT_ERROR}"
)


def _chain_violation(
    span: Span,
    toks: Sequence[VbaToken],
    start: int,
    steps: Sequence[FieldStep],
    states: Mapping[str, MemberState],
    lookup: IntegerConstantLookup,
) -> SubscriptHit | None:
    """The first subscript in a chain that the field cannot take, or a bound it lacks."""
    for step in steps:
        if not step.field.is_array:
            continue
        state = None if is_fixed_array_field(step.field) or not step.path else states.get(step.path)
        shape: FixedArrayBound | None
        if step.field.dims:
            shape = FixedArrayBound(name=step.display, dims=step.field.dims, origin="Dim")
        elif is_array_bounds(state):
            shape = dataclasses.replace(state, name=step.display)
        else:
            shape = None
        if step.open is not None:
            if state == "unallocated":
                assert step.close is not None
                return SubscriptHit(
                    span=Span(span.start + toks[start].start, span.start + toks[step.close].end),
                    message=f"'{step.display}' {_NO_ELEMENTS}",
                )
            hit = shape_subscript_violation(span, toks, shape, step.open, lookup) if shape is not None else None
            if hit is not None:
                return hit
            continue
        # Used whole: UBound(t.arr, 2), UBound(t.dyn).
        return _bound_violation(span, toks, start, step, "unallocated" if state == "unallocated" else shape)
    return None


def _bound_violation(
    span: Span,
    toks: Sequence[VbaToken],
    start: int,
    step: FieldStep,
    shape: str | FixedArrayBound | None,
) -> SubscriptHit | None:
    """`UBound(t.arr, 2)` past the field's dimensions, or UBound of a field with no elements."""
    callee = token_text(_at(toks, start - 2))
    after = _raw_at(toks, step.at + 1)
    if (
        callee not in ("ubound", "lbound")
        or _raw_at(toks, start - 1) != "("
        or not is_bare_or_vba_qualified_intrinsic_call(toks, start - 2)
        or after not in (")", ",")
        or not shape
    ):
        return None
    name = toks[start - 2].raw_text
    if shape == "unallocated":
        return SubscriptHit(
            span=Span(span.start + toks[start].start, span.start + toks[step.at].end),
            message=f"{name} reads the bounds of '{step.display}', which {_NO_ELEMENTS}",
        )
    assert isinstance(shape, FixedArrayBound)
    close = match_paren_from(toks, start - 1)
    dimension = toks[step.at + 2 : close] if after == "," and close > step.at + 2 else []
    value = (
        js_number(dimension[0].raw_text)
        if len(dimension) == 1 and dimension[0].kind is TokenKind.INTEGER_LITERAL
        else None
    )
    if value is None or 1 <= value <= len(shape.dims):
        return None
    return SubscriptHit(
        span=Span(span.start + dimension[0].start, span.start + dimension[0].end),
        message=(
            f"{name} asks for dimension {js_number_to_string(value)} of '{step.display}', which has "
            f"{pluralize_count(len(shape.dims), 'dimension')}. {_SUBSCRIPT_ERROR}"
        ),
    )


def _len_of_arrays(
    span: Span,
    toks: Sequence[VbaToken],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
    source_names: SourceNameScope,
    value_type: _ValueType = _no_value_type,
) -> list[tuple[Span, str, str]]:
    """`Len(a)` and `LenB(a)` of an array: a variable declared as one, or an
    array field used whole, parentheses or not. `VBA.Len(a)` is the library
    function instead, which takes the array as a Variant: a Byte array converts
    to a string and runs, and any other array raises 13 (measured in Excel
    16.0). Returns (span, message, rule)."""
    out: list[tuple[Span, str, str]] = []
    for i in range(len(toks) - 1):
        word = token_text(toks[i])
        if word not in ("len", "lenb") or toks[i + 1].raw_text != "(" or not is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        qualified = _raw_at(toks, i - 1) == "."
        if not qualified and runtime_callable_source_shadowed(toks[i].raw_text, source_names):
            continue
        close = match_paren_from(toks, i + 1)
        start = i + 2
        end = close - 1
        while start < end and toks[start].raw_text == "(" and match_paren_from(toks, start) == end:
            start += 1
            end -= 1
        if close < 0 or start > end:
            continue
        array = _array_named(toks, start, end, symbols, proc, types, subject)
        at = Span(span.start + toks[start].start, span.start + toks[end].end)
        if array is not None and not qualified:
            out.append(
                (
                    at,
                    f"'{array[0]}' is an array, which {toks[i].raw_text} cannot take. This is a VBE compile "
                    "error: Variable required - can't assign to this expression.",
                    "variableRequired",
                )
            )
        elif array is None and not qualified:
            type_ = _non_string_value_type(list(toks[start : end + 1]), symbols, proc, value_type)
            if type_:
                article = "an" if type_[0] in "AEIOU" else "a"
                text = "".join(tok.raw_text for tok in toks[start : end + 1])
                out.append(
                    (
                        at,
                        f"{toks[i].raw_text} of {article} {type_} reports a variable's storage size, and "
                        f"'{text}' is no variable. This is a VBE compile error: Variable required - can't "
                        "assign to this expression.",
                        "variableRequired",
                    )
                )
        elif array is not None and array[1] != "byte":
            out.append(
                (
                    at,
                    f"'{array[0]}' is an array, and not of Byte, so VBA.{toks[i].raw_text} cannot convert it "
                    "to a string. This will raise Run-time error '13': Type mismatch.",
                    "variantValueMisuse",
                )
            )
    return out


# VBA's functions that return one scalar type, not a Variant, as the VBE
# compiles `Len` of them (issue #475, measured in Excel 16.0): Asc is an
# Integer, Len and InStr a Long, Timer a Single.
_LIBRARY_RETURNS: Mapping[str, str] = {
    "asc": "integer",
    "ascw": "integer",
    "ascb": "integer",
    "len": "long",
    "lenb": "long",
    "instr": "long",
    "instrrev": "long",
    "timer": "single",
}


def _scope_symbols(symbols: ModuleSymbols, proc: ProcedureNode) -> list[VbaSymbol]:
    """The procedure's symbols, then the module's: `[...procedure children, ...root children]`."""
    from ...types.type_inference import procedure_symbol_for

    proc_sym = procedure_symbol_for(symbols, proc)
    return [
        *((proc_sym.children if proc_sym is not None else None) or []),
        *(symbols.root.children or []),
    ]


def _library_return_type(value: Sequence[VbaToken], symbols: ModuleSymbols, proc: ProcedureNode) -> str | None:
    """The type a VBA library value has, where the library fixes it: a call of
    one of _LIBRARY_RETURNS, Timer bare, a `$` function's String, and
    `Err.Number`'s Long. A name of the module's or the procedure's hides it."""
    lower = _lower_name(_at(value, 0))
    if (
        not lower
        or value[0].kind is TokenKind.BRACKETED_IDENTIFIER
        or any(child.name.lower() == lower for child in _scope_symbols(symbols, proc))
        or any(param.name.lower() == lower for param in proc.params)
    ):
        return None
    if len(value) == 3 and lower == "err" and value[1].raw_text == "." and token_text(value[2]) == "number":
        return "long"
    if len(value) == 1:
        return "single" if lower == "timer" else None
    # `Mid$(s, 1, 1)` lexes as Mid, a `$` and the arguments (issue #334).
    if value[1].raw_text == "$" and _raw_at(value, 2) == "(" and match_paren_from(value, 2) == len(value) - 1:
        # Only a String function takes a `$`; the registry knows Format$ as Format.
        fn = resolve_runtime_function(f"{lower}$") or resolve_runtime_function(lower)
        return "string" if fn is not None and fn.kind == "function" else None
    call = value[1].raw_text == "(" and match_paren_from(value, 1) == len(value) - 1
    return _LIBRARY_RETURNS.get(lower) if call else None


# VBA's functions whose return type the library declares as one scalar type: a
# conversion or Val.
_TYPED_RETURNS = frozenset({"cbool", "cbyte", "ccur", "cdate", "cdbl", "cint", "clng", "clnglng", "clngptr", "csng", "val"})


def _non_string_value_type(
    value: list[VbaToken], symbols: ModuleSymbols, proc: ProcedureNode, value_type: _ValueType
) -> str | None:
    """The type of a Len argument that is a value of known type other than
    String or Variant, and no variable, array element or field: a literal, an
    expression, a Const, a conversion or Val, or a project Function As Long
    (issue #368, measured in Excel 16.0). Len(Now), Len(Left(...)) and a
    Function As Variant or String compile."""
    name = token_name(value[0])
    named = variable_symbol_in(symbols, proc, name) if name else None
    whole_or_indexed = len(value) == 1 or (
        _raw_at(value, 1) == "(" and match_paren_from(value, 1) == len(value) - 1
    )
    if named is not None and named.kind is not VbaSymbolKind.CONSTANT and whole_or_indexed:
        return None
    word = token_text(value[0]) if len(value) == 1 else ""
    if word in ("true", "false"):
        return "Boolean"

    # `Len(c < v)`, `Len(i \ Round(1.5))`: an operation with a Variant operand
    # gives a Variant, and Len of a Variant compiles (issue #455).
    def settled(operand: list[VbaToken]) -> str | None:
        return _settled_type(operand, symbols, proc, value_type)

    if not settled(value):
        return None
    # `Len(i = 1)`: a comparison at the top level is a Boolean.
    depth = 0
    for tok in value:
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        if depth == 0 and (tok.raw_text in ("=", "<>", "<", ">", "<=", ">=") or token_text(tok) in ("like", "is")):
            return "Boolean"
    # `Const K = 5` then `Len(K)`: a Const is no variable.
    constant = (
        next((child for child in _scope_symbols(symbols, proc) if child.name.lower() == name.lower()), None)
        if len(value) == 1 and name
        else None
    )
    if constant is not None and constant.kind is VbaSymbolKind.CONSTANT:
        type_ = normalize_type(constant.as_type) or normalize_type(
            value_type([tok for tok in raw_expression_tokens(constant.default_raw or "") if tok.kind is not TokenKind.COMMENT])
        )
        return (
            type_[0].upper() + type_[1:]
            if type_ and type_ != "string" and type_ != "variant" and is_known_scalar_type(type_)
            else None
        )
    # A call names a project Function or a typed VBA function; any other library
    # function may return a Variant, as Now does.
    library = _library_return_type(value, symbols, proc)
    if library:
        return _scalar_name(library)
    call = bool(name) and _raw_at(value, 1) == "(" and match_paren_from(value, 1) == len(value) - 1
    if call and name is not None and name.lower() in _ARGUMENT_TYPED:
        return _scalar_name(settled(value))
    if (
        call
        and name is not None
        and name.lower() not in _TYPED_RETURNS
        and not any(
            child.kind is VbaSymbolKind.FUNCTION and child.name.lower() == name.lower()
            for child in (symbols.root.children or [])
        )
    ):
        return None
    if len(value) == 1 and name and named is None:
        return None
    # `\` and Mod give a whole number, never the Double value_type names.
    _operands, operators = _top_level_operands(value)
    if operators and all(op in ("\\", "mod") for op in operators):
        return _scalar_name(settled(value))
    # `Len(-d)`, `Len(Not i)`: a sign or Not keeps its operand's type.
    return _scalar_name(normalize_type(value_type(value)) or settled(value))


def _scalar_name(type_: str | None) -> str | None:
    """A type's display name, for a scalar type other than String or Variant."""
    if type_ and type_ != "string" and type_ != "variant" and is_known_scalar_type(type_):
        return type_[0].upper() + type_[1:]
    return None


# VBA functions whose result type follows their argument, measured in Excel 16.0
# with Len (issue #455): Sgn is an Integer and Sqr a Double whatever they take;
# Abs keeps its argument's type but gives a Variant for a Byte or Variant; Int and
# Fix keep a Single, Double or Currency and give a Variant for a Long or Integer
# variable.
_ARGUMENT_TYPED = frozenset({"abs", "int", "fix", "sgn", "sqr"})

_COMPARISONS = frozenset({"=", "<>", "<", ">", "<=", ">=", "like", "is"})
_BINARY_OPERATORS = frozenset(
    {*_COMPARISONS, "+", "-", "*", "/", "\\", "^", "&", "mod", "and", "or", "xor", "eqv", "imp"}
)
_INTEGRAL_RANK: Mapping[str, int] = {"boolean": 1, "byte": 1, "integer": 2, "long": 3, "longlong": 4}
# The type a number literal's suffix gives it.
_LITERAL_SUFFIX_TYPES: Mapping[str, str] = {
    "#": "double",
    "!": "single",
    "@": "currency",
    "%": "integer",
    "&": "long",
    "^": "longlong",
}


def _top_level_operands(value: Sequence[VbaToken]) -> tuple[list[list[VbaToken]], list[str]]:
    """The operands and operators at the top level of an expression, outer
    parentheses and unary signs left on the operands."""
    operands: list[list[VbaToken]] = [[]]
    operators: list[str] = []
    depth = 0
    for tok in value:
        text = token_text(tok)
        current = operands[-1]
        depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
        last = current[-1] if current else None
        after_operand = (
            last is not None
            and not (last.kind is TokenKind.OPERATOR and last.raw_text != ")")
            and token_text(last) != "not"
        )
        if depth == 0 and after_operand and text in _BINARY_OPERATORS:
            operators.append(text)
            operands.append([])
        else:
            current.append(tok)
    return operands, operators


def _settled_type(
    toks: list[VbaToken], symbols: ModuleSymbols, proc: ProcedureNode, value_type: _ValueType
) -> str | None:
    """The type of a value, lowercased, where VBA gives it a type other than
    Variant, or None where it is a Variant or unsettled: a literal, a Const, a
    declared variable, a typed VBA function or project Function, a unary sign or
    Not, and an operation none of whose operands may be a Variant (issue #455,
    measured in Excel 16.0)."""
    value = toks
    while len(value) > 2 and value[0].raw_text == "(" and match_paren_from(value, 0) == len(value) - 1:
        value = value[1:-1]
    operands, operators = _top_level_operands(value)
    if operators:
        # `o Is Nothing` is a Boolean whatever the operands.
        if len(operators) == 1 and operators[0] == "is":
            return "boolean"
        types = [_settled_type(operand, symbols, proc, value_type) if operand else None for operand in operands]
        if any(not type_ for type_ in types):
            return None
        if any(op in _COMPARISONS for op in operators):
            return "boolean"
        if "&" in operators:
            return "string"
        if all(op in ("\\", "mod") for op in operators):
            rank = max(_INTEGRAL_RANK.get(type_ or "", 3) for type_ in types)
            return "longlong" if rank >= 4 else "long" if rank == 3 else "integer"
        return normalize_type(value_type(value)) or types[0]
    if not value:
        return None
    head = token_text(value[0])
    if len(value) > 1 and head in ("-", "+", "not"):
        return _settled_type(value[1:], symbols, proc, value_type)
    if len(value) == 1:
        tok = value[0]
        if tok.kind is TokenKind.STRING_LITERAL:
            return "string"
        if tok.kind is TokenKind.DATE_LITERAL:
            return "date"
        if tok.kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL):
            suffixed = _LITERAL_SUFFIX_TYPES.get(tok.raw_text[-1:])
            if suffixed is not None:
                return suffixed
            whole = js_number(tok.raw_text) if tok.kind is TokenKind.INTEGER_LITERAL else math.nan
            if math.isnan(whole):
                return "double"
            return "integer" if abs(whole) <= 32767 else "long" if abs(whole) <= 2147483647 else "double"
        if head in ("true", "false"):
            return "boolean"
    name = token_name(value[0])
    if not name:
        return None
    lower = name.lower()
    library = _library_return_type(value, symbols, proc)
    if library:
        return library
    call = len(value) > 2 and value[1].raw_text == "(" and match_paren_from(value, 1) == len(value) - 1
    if len(value) == 1:
        symbol = next((child for child in _scope_symbols(symbols, proc) if child.name.lower() == lower), None)
        if symbol is not None and symbol.kind is VbaSymbolKind.CONSTANT:
            return normalize_type(symbol.as_type) or normalize_type(
                value_type([t for t in raw_expression_tokens(symbol.default_raw or "") if t.kind is not TokenKind.COMMENT])
            )
        variable = variable_symbol_in(symbols, proc, lower)
        type_ = normalize_type(variable.as_type if variable is not None else None)
        return type_ if type_ and type_ != "variant" else None
    if not call:
        return None
    if lower == "sgn":
        return "integer"
    if lower == "sqr":
        return "double"
    if lower in ("abs", "int", "fix"):
        argument = _settled_type(value[2:-1], symbols, proc, value_type)
        # Fix(2) is refused where Fix(i) compiles: a literal keeps its type.
        literal = len(value) == 4 and value[2].kind in (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL)
        kept = argument != "byte" if lower == "abs" else literal or argument in ("single", "double", "currency")
        if not kept:
            return None
        return "double" if lower == "abs" and argument == "string" else argument
    if lower in _TYPED_RETURNS:
        return normalize_type(value_type(value))
    fn = next(
        (
            child
            for child in (symbols.root.children or [])
            if child.kind is VbaSymbolKind.FUNCTION and child.name.lower() == lower
        ),
        None,
    )
    type_ = normalize_type(fn.as_type if fn is not None else None)
    return type_ if type_ and type_ != "variant" else None


_EMPTY_PARENS_SUFFIX_RE = re.compile(f"\\([{JS_WHITESPACE}]*\\)[{JS_WHITESPACE}]*\\Z")


def _array_named(
    toks: Sequence[VbaToken],
    start: int,
    end: int,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
) -> tuple[str, str | None] | None:
    """The array the tokens from `start` to `end` name whole, a variable or a
    field, and its element type, lowercased: (display, element type)."""
    if start == end:
        lower = _lower_name(toks[start])
        variable = variable_symbol_in(symbols, proc, lower) if lower else None
        if variable is not None and variable.is_array and not variable.param_array:
            as_type = _EMPTY_PARENS_SUFFIX_RE.sub("", variable.as_type, count=1) if variable.as_type is not None else None
            return (toks[start].raw_text, type_key(as_type))
        return None
    root = type_root_at(toks, start, symbols, proc, types, subject)
    steps = field_chain(toks, root, types) if root is not None else []
    last = steps[-1] if steps else None
    if last is not None and last.field.is_array and last.open is None and last.at == end:
        return (last.display, last.field.type)
    return None
