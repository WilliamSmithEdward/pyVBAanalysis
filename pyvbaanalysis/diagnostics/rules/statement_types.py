"""Rule family: statements the VBE refuses by their form or by the type of
what they name (XLIDE issue #213); statement_forms.py holds the #125 family.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/statementTypes.ts.
Each verdict is a full project compile in 64-bit Excel 16.0 (build 20326,
2026-09-30):

- for-variable-in-use: a For or For Each inside another on the same
  control variable, "For control variable already in use".
- for-counter-type: a For counter that is a String, Boolean, object or
  user-defined type, "Type mismatch". Variant, Date and the numbers take.
- for-each-source-type: For Each over an array of a user-defined type or
  of fixed-length strings, fixed or dynamic.
- statement-before-first-case: a statement or label between Select Case
  and its first Case; a comment is fine.
- return-with-value: `Return 5`, "Syntax error". A bare Return is GoSub's.
- with-scalar-target: With on a number, a string or a scalar variable,
  "With object must be user-defined type, Object, or Variant".
- set-requires-object: Set of a user-defined type variable.
- paramarray-passing-mode: ByVal or ByRef on a ParamArray, "Expected:
  identifier".
- udt-value-mismatch and udt-variant-coercion: a user-defined type where
  a single value is needed, and one handed to a Variant.

Issue #216, measured the same way:

- const-value-not-constant and array-bound-not-constant: a Const value or
  a Dim bound that names a variable, "Constant expression required".
  ReDim takes one.
- variable-required: a Const as a For counter or a Mid target, "Variable
  required - can't assign to this expression".
- type-suffix-mismatch: `n% = 2` with n As Long, "Type-declaration
  character does not match declared data type". The suffix of the
  declared type is fine.
- named-argument-not-allowed: InStr, Len, StrComp, Abs, Int, Fix, Sgn and
  the C-conversions but CDec take no named arguments, "Syntax error".

Issue #253, measured the same way:

- lset-type-mismatch: LSet between two different user-defined types,
  either of which holds a variable-length String, a dynamic array, an
  object or a Variant, "Type mismatch". Two Types of fixed-size
  members, and two values of one Type, take it.
- udt-variant-coercion: a user-defined type passed to a method of a
  Collection, an Object or a Variant, `c.Add t`, `o.Add t`.
"""

from __future__ import annotations

import dataclasses
import re
import unicodedata
from collections.abc import Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Literal

from ...completion.type_completion import ProjectTypeName
from ...conditional import ConditionalActivityTracker
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    AssignmentNode,
    BodyNode,
    CallNode,
    ConditionalDirectiveNode,
    DeclareNode,
    EnumNode,
    ForBlockNode,
    IfBlockNode,
    ModuleNode,
    ProcedureNode,
    SelectBlockNode,
    Span,
    StatementNode,
    TypeNode,
    VariableGroupNode,
    WithBlockNode,
    iter_body_nodes_in_context,
)
from ...runtime import resolve_runtime_function
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import (
    ModuleSymbols,
    VbaProjectTypeKind,
    VbaSymbol,
    VbaSymbolKind,
)
from ...types.type_inference import procedure_symbol_for, source_identifier_binding
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn
from ..type_fields import ModuleTypes, is_fixed_array_field, module_types
from ..walker import (
    absolute_span,
    active_module_members,
    is_inactive_node,
    match_paren_from,
    raw_expression_tokens,
    statement_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# What a variable holds, as far as these rules care.
_Category = Literal["variant", "number", "date", "string", "boolean", "object", "udt", "class"]

_NUMBER_TYPES = frozenset(
    {"byte", "integer", "long", "longlong", "longptr", "currency", "single", "double", "decimal"}
)

# Words that join two operands, so a user-defined type beside one is used as a value.
_OPERATOR_WORDS = frozenset({"and", "or", "xor", "eqv", "imp", "mod", "like", "not", "is"})

# `/^[ \t]*Def(?:...)\b/im`: JavaScript's `^` under the m flag also follows a lone
# CR, U+2028 and U+2029, and its i flag folds ASCII only.
_DEFTYPE_LINE_RE = re.compile(
    "(?:^|(?<=[\\r\\u2028\\u2029]))[ \\t]*Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec"
    "|Date|Str|Obj|Var)\\b",
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_LINE_BREAK_RE = re.compile(r"\r|\n")


@dataclass(slots=True)
class _Context:
    source: str
    symbols: ModuleSymbols
    project_visible_symbols: Sequence[VbaSymbol] | None
    activity: ConditionalActivityTracker | None
    push: PushFn
    udts: AbstractSet[str]
    enums: AbstractSet[str]
    # Whether an untyped variable is a Variant, which a Deftype line changes.
    untyped_is_variant: bool
    # The Types this module declares, with their fields.
    types: ModuleTypes


def check_statement_types(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    project_types: Sequence[ProjectTypeName] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    udts: set[str] = set()
    enums: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, TypeNode):
            udts.add(member.name.lower())
        elif isinstance(member, EnumNode):
            enums.add(member.name.lower())
    for type_ in project_types or []:
        if type_.kind is VbaProjectTypeKind.USER_TYPE:
            udts.add(type_.name.lower())
        elif type_.kind is VbaProjectTypeKind.ENUM:
            enums.add(type_.name.lower())
    ctx = _Context(
        source=source,
        symbols=symbols,
        project_visible_symbols=project_visible_symbols,
        activity=activity,
        push=push,
        udts=udts,
        enums=enums,
        untyped_is_variant=_DEFTYPE_LINE_RE.search(source) is None,
        types=module_types(source, mod, activity),
    )
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            _check_declaration_group(ctx, None, member)
        if isinstance(member, (ProcedureNode, DeclareNode)):
            for param in member.params:
                if param.param_array and (param.by_val or param.by_ref):
                    push(
                        "paramArrayPassingMode",
                        "A ParamArray is always passed ByRef, and takes neither ByVal nor ByRef: "
                        f"'{param.name}'. This is a VBE compile error: Expected: identifier.",
                        param.name_span if param.name_span is not None else param.span,
                    )
        if isinstance(member, ProcedureNode):
            proc_sym = procedure_symbol_for(symbols, member)
            _walk_body(ctx, member, proc_sym, member.body)


def _variable_named(ctx: _Context, proc_sym: VbaSymbol | None, name: str) -> VbaSymbol | None:
    """The single variable a bare name binds to, or None."""
    binding = source_identifier_binding(
        ctx.symbols, proc_sym, ctx.project_visible_symbols, name, BareIdentifierContext.EXPRESSION
    )
    if (
        binding.scope is BareIdentifierResolutionScope.UNRESOLVED
        or binding.scope is BareIdentifierResolutionScope.AMBIGUOUS
        or len(binding.definitions) != 1
    ):
        return None
    definition = binding.definitions[0]
    return (
        definition
        if definition.kind
        in (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.PARAMETER)
        else None
    )


def _scalar_function_named(ctx: _Context, proc_sym: VbaSymbol | None, name: str) -> str | None:
    """The scalar type a Function of the module or project returns, when the name is one."""
    binding = source_identifier_binding(
        ctx.symbols, proc_sym, ctx.project_visible_symbols, name, BareIdentifierContext.EXPRESSION
    )
    # VBA's own `CStr(1)`, where no name of the project hides it (XLIDE issue #612,
    # measured in Excel 16.0).
    if binding.scope is BareIdentifierResolutionScope.UNRESOLVED:
        runtime = resolve_runtime_function(name)
        runtime_type = (
            normalize_type(runtime.returns)
            if runtime is not None and runtime.kind == "function"
            else None
        )
        return (
            runtime.returns
            if runtime is not None
            and runtime_type
            and runtime_type != "variant"
            and is_known_scalar_type(runtime_type)
            else None
        )
    if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS or len(binding.definitions) != 1:
        return None
    definition = binding.definitions[0]
    type_ = normalize_type(definition.as_type)
    return (
        definition.as_type
        if (
            definition.kind is VbaSymbolKind.FUNCTION
            or (definition.kind is VbaSymbolKind.DECLARE and definition.declare_kind == "Function")
        )
        and not definition.is_array
        and type_
        and type_ != "variant"
        and type_ != "object"
        and is_known_scalar_type(type_)
        else None
    )


def _category(ctx: _Context, symbol: VbaSymbol | None) -> _Category | None:
    if symbol is None or symbol.is_array:
        return None
    type_ = normalize_type(symbol.as_type)
    if not type_:
        return "variant" if ctx.untyped_is_variant else None
    if type_ == "variant":
        return "variant"
    if type_ in _NUMBER_TYPES or type_ in ctx.enums:
        return "number"
    if type_ == "date":
        return "date"
    if type_ == "string":
        return "string"
    if type_ == "boolean":
        return "boolean"
    if type_ == "object":
        return "object"
    if type_ in ctx.udts:
        return "udt"
    # A name no scalar type or Enum claims is a class: Collection, Range, a
    # class module.
    return None if is_known_scalar_type(type_) else "class"


def _walk_body(
    ctx: _Context,
    member: ProcedureNode,
    proc_sym: VbaSymbol | None,
    body: Sequence[BodyNode],
) -> None:
    """walkBody, on an explicit stack: each block's body is walked with the For
    control variables of the blocks around it."""
    # The counters a node's own body is walked with, set as the node is visited
    # and read when its body is entered, which follows at once.
    entered: list[tuple[str, ...]] = [()]

    def enter(node: BodyNode, counters: tuple[str, ...]) -> tuple[str, ...]:
        return entered[0]

    def skip(node: BodyNode) -> bool:
        return is_inactive_node(ctx.activity, node)

    start: tuple[str, ...] = ()
    for node, counters in iter_body_nodes_in_context(body, start, enter, skip):
        inner = counters
        if isinstance(node, VariableGroupNode):
            _check_declaration_group(ctx, proc_sym, node)
        elif isinstance(node, ForBlockNode):
            inner = _check_for(ctx, proc_sym, node, counters)
        elif isinstance(node, SelectBlockNode):
            _check_select_head(ctx, node.body)
        elif isinstance(node, WithBlockNode):
            _check_with_target(ctx, proc_sym, node.span)
        elif isinstance(node, IfBlockNode):
            for branch in node.branches:
                _check_condition(ctx, proc_sym, branch.condition_raw, branch.header_span)
        elif isinstance(node, (StatementNode, AssignmentNode, CallNode)):
            _check_statement(ctx, proc_sym, node.span)
        entered[0] = inner


def _is_unicode_letter(ch: str) -> bool:
    """`\\p{L}`."""
    return unicodedata.category(ch).startswith("L")


def _is_counter_name(name: str) -> bool:
    """`/^[\\p{L}_][\\p{L}\\p{N}_]*$/u`."""
    if not name or not (name[0] == "_" or _is_unicode_letter(name[0])):
        return False
    return all(ch == "_" or unicodedata.category(ch)[0] in ("L", "N") for ch in name[1:])


def _check_for(
    ctx: _Context,
    proc_sym: VbaSymbol | None,
    node: ForBlockNode,
    counters: tuple[str, ...],
) -> tuple[str, ...]:
    name = js_trim(node.control_variable) if node.control_variable is not None else None
    if not name or not _is_counter_name(name) or node.control_variable_span is None:
        return counters
    key = name.lower()
    if _is_constant_name(ctx, proc_sym, name):
        _variable_required(ctx, name, node.control_variable_span)
    if key in counters:
        ctx.push(
            "forVariableInUse",
            f"'{name}' is already the control variable of an enclosing For. This is a VBE compile "
            "error: For control variable already in use.",
            node.control_variable_span,
        )
    variable = _variable_named(ctx, proc_sym, name)
    if not node.each:
        kind = _category(ctx, variable)
        if kind in ("string", "boolean", "object", "udt", "class") and variable is not None:
            ctx.push(
                "forCounterType",
                f"A For counter must be a number, a Date or a Variant, and '{name}' is declared As "
                f"{variable.as_type}. This is a VBE compile error: Type mismatch.",
                node.control_variable_span,
            )
    elif node.source_expression and node.source_expression_span is not None:
        source_tokens = raw_expression_tokens(node.source_expression)
        source_name = token_name(source_tokens[0]) if len(source_tokens) == 1 else None
        array = _variable_named(ctx, proc_sym, source_name) if source_name else None
        element = normalize_type(array.as_type if array is not None else None)
        if (
            array is not None
            and array.is_array
            and (array.fixed_length is not None or (element is not None and element in ctx.udts))
        ):
            what = (
                "fixed-length strings"
                if array.fixed_length is not None
                else f"the user-defined type {array.as_type}"
            )
            ctx.push(
                "forEachSourceType",
                f"For Each cannot walk '{source_name}', an array of {what}. This is a VBE compile "
                "error: For Each may not be used on array of user-defined type or fixed-length "
                "strings.",
                node.source_expression_span,
            )
    return (*counters, key)


def _check_select_head(ctx: _Context, body: Sequence[BodyNode]) -> None:
    """A statement or label between `Select Case` and its first `Case`."""
    for node in body:
        if is_inactive_node(ctx.activity, node) or isinstance(node, ConditionalDirectiveNode):
            continue
        toks = statement_tokens(ctx.source, node.span)
        if len(toks) == 0:
            continue
        if token_text(toks[0]) == "case":
            return
        ctx.push(
            "statementBeforeFirstCase",
            "Nothing but comments may come between Select Case and its first Case. This is a VBE "
            "compile error: Statements and labels invalid between Select Case and first Case.",
            Span(node.span.start + toks[0].start, node.span.start + toks[-1].end),
        )
        return


def _header_tokens(source: str, span: Span) -> list[VbaToken]:
    """The `With` line's tokens, up to the end of that physical line."""
    found = _LINE_BREAK_RE.search(source, span.start, span.end)
    line_end = found.start() - span.start if found is not None else -1
    toks = statement_tokens(
        source, span if line_end < 0 else Span(span.start, span.start + line_end)
    )
    end = next(
        (
            k
            for k, tok in enumerate(toks)
            if tok.kind is TokenKind.COMMENT or tok.kind is TokenKind.COLON
        ),
        -1,
    )
    return toks if end < 0 else toks[:end]


def _check_with_target(ctx: _Context, proc_sym: VbaSymbol | None, span: Span) -> None:
    toks = _header_tokens(ctx.source, span)
    if token_text(_at(toks, 0)) != "with":
        return
    # `With a(1)` on an array of Long, `With TakeL(1)` on a Function that
    # returns one (XLIDE issue #325, measured in Excel 16.0).
    if (
        len(toks) > 3
        and toks[2].raw_text == "("
        and match_paren_from(toks, 2) == len(toks) - 1
        and token_name(toks[1])
    ):
        name = toks[1].raw_text
        array = _variable_named(ctx, proc_sym, name)
        element = (
            _category(ctx, dataclasses.replace(array, is_array=False))
            if array is not None and array.is_array
            else None
        )
        fn = None if array is not None else _scalar_function_named(ctx, proc_sym, name)
        if element in ("number", "string", "boolean", "date") or fn:
            target_text = "".join(tok.raw_text for tok in toks[1:])
            what = f"a {fn}, what the Function returns" if fn else "an element of a scalar array"
            ctx.push(
                "withScalarTarget",
                f"With needs an object, a user-defined type or a Variant, and '{target_text}' is "
                f"{what}. This is a VBE compile error: With object must be user-defined type, "
                "Object, or Variant.",
                Span(span.start + toks[1].start, span.start + toks[-1].end),
            )
        return
    if len(toks) != 2:
        return
    target = toks[1]
    literal = target.kind in (
        TokenKind.STRING_LITERAL,
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.DATE_LITERAL,
    )
    target_name = None if literal else token_name(target)
    kind = _category(ctx, _variable_named(ctx, proc_sym, target_name)) if target_name else None
    if literal or kind in ("number", "string", "boolean", "date"):
        ctx.push(
            "withScalarTarget",
            f"With needs an object, a user-defined type or a Variant, and '{target.raw_text}' is "
            f"{'a literal' if literal else 'a scalar'}. This is a VBE compile error: With object "
            "must be user-defined type, Object, or Variant.",
            absolute_span(span, target),
        )


def _check_condition(
    ctx: _Context, proc_sym: VbaSymbol | None, condition_raw: str | None, header_span: Span
) -> None:
    toks = raw_expression_tokens(condition_raw) if condition_raw else []
    name = token_name(toks[0]) if len(toks) == 1 else None
    if name and condition_raw and _category(ctx, _variable_named(ctx, proc_sym, name)) == "udt":
        at = ctx.source.find(js_trim(condition_raw), header_span.start)
        _udt_mismatch(ctx, name, header_span if at < 0 else Span(at, at + len(name)))


def _udt_mismatch(ctx: _Context, name: str, span: Span) -> None:
    ctx.push(
        "udtValueMismatch",
        f"'{name}' is a user-defined type, which cannot be used as a single value here. This is a "
        "VBE compile error: Type mismatch.",
        span,
    )


def _udt_coercion(ctx: _Context, name: str, span: Span) -> None:
    ctx.push(
        "udtVariantCoercion",
        f"'{name}' is a user-defined type, and a Variant cannot hold one declared in a standard "
        "module or a private class. This is a VBE compile error: Only user-defined types defined "
        "in public object modules can be coerced to or from a variant or passed to late-bound "
        "functions.",
        span,
    )


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _article(type_name: str) -> str:
    """`/^[aeiou]/i.test(typeName) ? 'an' : 'a'`."""
    return "an" if type_name[:1] != "" and type_name[0] in "aeiouAEIOU" else "a"


_LITERAL_KINDS = (
    TokenKind.STRING_LITERAL,
    TokenKind.INTEGER_LITERAL,
    TokenKind.FLOAT_LITERAL,
    TokenKind.DATE_LITERAL,
)


def _check_statement(ctx: _Context, proc_sym: VbaSymbol | None, span: Span) -> None:
    toks = statement_tokens_after_leading_label(ctx.source, span)
    if len(toks) == 0:
        return
    _check_type_suffixes(ctx, proc_sym, toks, span)
    _check_named_arguments(ctx, proc_sym, toks, span)
    _check_mid_target(ctx, proc_sym, toks, span)
    first = token_text(toks[0])
    if first == "return" and len(toks) > 1 and toks[1].kind is not TokenKind.COLON:
        ctx.push(
            "returnWithValue",
            "Return takes no value in VBA: it goes back from a GoSub. Assign the function name "
            "instead. This is a VBE compile error: Syntax error.",
            Span(span.start + toks[0].start, span.start + toks[-1].end),
        )
        return

    def udt_at(i: int) -> str | None:
        tok = _at(toks, i)
        name = token_name(tok) if tok is not None else None
        if (
            not name
            or _raw_at(toks, i - 1) == "."
            or _raw_at(toks, i + 1) == "."
            or _raw_at(toks, i + 1) == "("
        ):
            return None
        return name if _category(ctx, _variable_named(ctx, proc_sym, name)) == "udt" else None

    def at(i: int) -> Span:
        return absolute_span(span, toks[i])

    if first == "set":
        name = udt_at(1)
        if name and _raw_at(toks, 2) == "=":
            ctx.push(
                "setRequiresObject",
                f"'{name}' is a user-defined type, not an object, so Set cannot assign it. This is "
                "a VBE compile error: Object required.",
                at(1),
            )
        return
    if first == "debug" and _raw_at(toks, 1) == "." and token_text(_at(toks, 2)) == "print":
        for i in range(3, len(toks)):
            name = udt_at(i)
            if (
                name
                and _is_argument_boundary(_at(toks, i - 1))
                and _is_argument_boundary(_at(toks, i + 1))
            ):
                _udt_mismatch(ctx, name, at(i))
    # `LSet n = 5` on a Long, `RSet a = b` on a Type (XLIDE issue #451, measured
    # in Excel 16.0). LSet takes a String or a Type, and runs on a Variant;
    # RSet takes a String or a Variant only.
    if (first == "lset" or first == "rset") and _raw_at(toks, 2) == "=":
        target_name = token_name(_at(toks, 1))
        target = _variable_named(ctx, proc_sym, target_name if target_name is not None else "")
        kind = _category(ctx, target)
        refused = kind in ("number", "boolean", "date") or (first == "rset" and kind == "udt")
        if target is not None and refused and not target.is_array:
            as_type = target.as_type if target.as_type is not None else ""
            error = (
                "LSet allowed only on strings and user-defined types"
                if first == "lset"
                else "RSet allowed only on strings"
            )
            ctx.push(
                "lsetTypeMismatch",
                f"'{toks[1].raw_text}' is {_article(as_type)} {target.as_type}, which "
                f"{toks[0].raw_text} cannot fill. This is a VBE compile error: {error}.",
                Span(span.start + toks[0].start, span.start + toks[1].end),
            )
            return
    if first == "lset" and len(toks) == 4 and toks[2].raw_text == "=":
        _check_lset(ctx, proc_sym, toks, span)
        return
    _check_method_arguments(ctx, proc_sym, toks, span)
    if first == "msgbox":
        start = 2 if _raw_at(toks, 1) == "(" else 1
        name = udt_at(start)
        if name and _is_argument_boundary(_at(toks, start + 1)):
            _udt_coercion(ctx, name, at(start))
        return

    # `a = b`: the target and the value's kinds, when each is a bare name or a
    # literal.
    eq = 2 if first == "let" else 1
    target_name = token_name(_at(toks, eq - 1))
    if _raw_at(toks, eq) == "=" and len(toks) == eq + 2 and target_name:
        target = _variable_named(ctx, proc_sym, target_name)
        target_kind = _category(ctx, target)
        value = toks[eq + 1]
        value_name = token_name(value)
        value_symbol = _variable_named(ctx, proc_sym, value_name) if value_name else None
        value_kind = _category(ctx, value_symbol)
        literal = value.kind in _LITERAL_KINDS
        if target_kind == "udt":
            other_type = (
                value_kind == "udt"
                and value_symbol is not None
                and target is not None
                and normalize_type(value_symbol.as_type) != normalize_type(target.as_type)
            )
            if literal or other_type or value_kind in ("number", "string", "boolean", "date"):
                _udt_mismatch(ctx, target_name, at(eq - 1))
        elif value_kind == "udt" and value_name is not None:
            if target_kind == "variant":
                _udt_coercion(ctx, value_name, at(eq + 1))
            elif target_kind in ("number", "string", "boolean", "date"):
                _udt_mismatch(ctx, value_name, at(eq + 1))
        return

    # A user-defined type beside an operator is used as a value: `t + 1`,
    # `t = u` in an expression. The assignment's own `=` is not one, however
    # long its target: `parents(depth) = cInfo`, `.Items(0).Id = rec`.
    assignment_eq = _first_top_level_equals(toks) if _starts_assignment(toks) else -1
    for i in range(len(toks)):
        name = udt_at(i)
        if not name:
            continue
        before = _at(toks, i - 1)
        after = _at(toks, i + 1)
        operator_before = before is not None and i - 1 != assignment_eq and _is_operator(before)
        operator_after = after is not None and i + 1 != assignment_eq and _is_operator(after)
        if operator_before or operator_after:
            _udt_mismatch(ctx, name, at(i))


def _starts_assignment(toks: Sequence[VbaToken]) -> bool:
    """Whether a statement is an assignment by its first token: a name, a
    `.member`, Let, Set, or LSet and RSet, which copy one user-defined type
    into another."""
    first = _at(toks, 0)
    if first is None:
        return False
    word = token_text(first)
    return (
        first.raw_text == "."
        or word in ("let", "set", "lset", "rset")
        or first.kind is TokenKind.IDENTIFIER
        or first.kind is TokenKind.BRACKETED_IDENTIFIER
    )


def _first_top_level_equals(toks: Sequence[VbaToken]) -> int:
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0 and raw == "=":
            return i
    return -1


def _is_operator(tok: VbaToken) -> bool:
    # `Item:=t` names an argument, which udt-variant-coercion reads.
    if tok.kind is TokenKind.OPERATOR:
        return tok.raw_text != "." and tok.raw_text != "!" and tok.raw_text != ":="
    return token_text(tok) in _OPERATOR_WORDS


def _check_lset(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    """`LSet a = b` between two Types (XLIDE issue #253): refused when they differ and
    either holds a member that is not of fixed size."""
    target_name = token_name(toks[1])
    value_name = token_name(toks[3])
    target = _variable_named(ctx, proc_sym, target_name if target_name is not None else "")
    value = _variable_named(ctx, proc_sym, value_name if value_name is not None else "")
    target_type = normalize_type(target.as_type if target is not None else None)
    value_type = normalize_type(value.as_type if value is not None else None)
    if (
        _category(ctx, target) != "udt"
        or _category(ctx, value) != "udt"
        or not target_type
        or not value_type
        or target_type == value_type
    ):
        return
    held = _variable_size_member(ctx, target_type, 0)
    if held is None:
        held = _variable_size_member(ctx, value_type, 0)
    if not held or target is None or value is None:
        return
    ctx.push(
        "lsetTypeMismatch",
        f"LSet copies '{toks[3].raw_text}' ({value.as_type}) into '{toks[1].raw_text}' "
        f"({target.as_type}), two different user-defined types, and the member {held}. LSet "
        "copies between two types only when every member is of fixed size. This is a VBE "
        "compile error: Type mismatch.",
        Span(span.start + toks[0].start, span.start + toks[3].end),
    )


def _variable_size_member(ctx: _Context, type_: str, depth: int) -> str | None:
    """A member of a module Type that is not of fixed size, described; None when
    every member is, or one is unknown."""
    fields = ctx.types.get(type_)
    if fields is None or depth > 8:
        return None
    for field in fields.values():
        field_type = normalize_type(field.type)
        if field.is_array and not is_fixed_array_field(field):
            return f"'{field.name}' is a dynamic array"
        if not field_type or field_type == "variant":
            return f"'{field.name}' is a Variant"
        if field_type == "string" and field.fixed_length is None:
            return f"'{field.name}' is a variable-length String"
        if field_type in ctx.types:
            nested = _variable_size_member(ctx, field_type, depth + 1)
            if nested:
                return nested
        elif field_type == "object" or field_type == "collection":
            return f"'{field.name}' is an object"
    return None


# VBA's functions measured to refuse a user-defined type for their Variant
# parameter (XLIDE issue #253).
_VARIANT_FUNCTIONS = frozenset(
    {
        "array", "choose", "cvar", "format", "iif", "isarray", "isdate", "isempty", "iserror",
        "ismissing", "isnull", "isnumeric", "isobject", "typename", "vartype",
    }
)  # fmt: skip


def _check_qualified_library_arguments(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    """`VBA.TypeName(t)`: the qualified forms of VBA's functions refuse a Type as
    the bare ones do; argument-shape-mismatch reads the bare ones."""
    for i in range(len(toks) - 3):
        if (
            token_text(toks[i]) != "vba"
            or _raw_at(toks, i - 1) == "."
            or toks[i + 1].raw_text != "."
            or token_text(toks[i + 2]) not in _VARIANT_FUNCTIONS
            or toks[i + 3].raw_text != "("
        ):
            continue
        depth = 0
        for j in range(i + 4, len(toks)):
            raw = toks[j].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                if depth == 0:
                    break
                depth -= 1
            else:
                name = token_name(toks[j])
                if (
                    depth == 0
                    and name
                    and toks[j - 1].raw_text in (",", "(")
                    and (_raw_at(toks, j + 1) or "") in (",", ")")
                    and _category(ctx, _variable_named(ctx, proc_sym, name)) == "udt"
                ):
                    _udt_coercion(ctx, toks[j].raw_text, absolute_span(span, toks[j]))


_RECEIVER_TYPES = frozenset({"collection", "object", "variant"})


def _check_method_arguments(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    """`c.Add t`, `o.Add Item:=t`, `Call c.Add(t)`: a user-defined type passed to
    a method of a Collection, an Object or a Variant (XLIDE issue #253)."""
    _check_qualified_library_arguments(ctx, proc_sym, toks, span)
    at = 1 if token_text(_at(toks, 0)) == "call" else 0
    receiver_name = token_name(_at(toks, at))
    receiver = _variable_named(ctx, proc_sym, receiver_name) if receiver_name else None
    receiver_type = normalize_type(receiver.as_type if receiver is not None else None)
    if receiver_type is None and ctx.untyped_is_variant:
        receiver_type = "variant"
    if (
        receiver is None
        or receiver.is_array
        or (receiver_type or "") not in _RECEIVER_TYPES
        or _raw_at(toks, at + 1) != "."
        or not token_name(_at(toks, at + 2))
    ):
        return
    start = at + 3
    end = len(toks)
    if _raw_at(toks, start) == "(" and _raw_at(toks, len(toks) - 1) == ")":
        start += 1
        end -= 1
    depth = 0
    slot_start = start
    for i in range(start, end + 1):
        raw = _raw_at(toks, i)
        if i < end and raw != ",":
            depth += 1 if raw == "(" else -1 if raw == ")" else 0
            continue
        if depth > 0:
            continue
        slot = list(toks[slot_start:i]) if slot_start <= i else []
        slot_start = i + 1
        value = slot[2:] if _raw_at(slot, 1) == ":=" else slot
        name = token_name(value[0]) if len(value) == 1 else None
        if name and _category(ctx, _variable_named(ctx, proc_sym, name)) == "udt":
            _udt_coercion(ctx, name, absolute_span(span, value[0]))


def _is_argument_boundary(tok: VbaToken | None) -> bool:
    """Whether a token ends or begins a Print or MsgBox argument."""
    return (
        tok is None
        or tok.raw_text in (",", ";", ")", "(")
        or tok.kind is TokenKind.COLON
        or token_text(tok) == "print"
    )


def _is_constant_name(ctx: _Context, proc_sym: VbaSymbol | None, name: str) -> bool:
    """A name that binds to a Const or an Enum member."""
    binding = source_identifier_binding(
        ctx.symbols, proc_sym, ctx.project_visible_symbols, name, BareIdentifierContext.EXPRESSION
    )
    return (
        binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS
        and binding.scope is not BareIdentifierResolutionScope.UNRESOLVED
        and len(binding.definitions) > 0
        and all(
            definition.kind in (VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM_MEMBER)
            for definition in binding.definitions
        )
    )


def _variable_required(ctx: _Context, name: str, span: Span) -> None:
    ctx.push(
        "variableRequired",
        f"'{name}' is a constant, and a constant cannot be assigned to here. This is a VBE "
        "compile error: Variable required - can't assign to this expression.",
        span,
    )


def _check_declaration_group(
    ctx: _Context, proc_sym: VbaSymbol | None, group: VariableGroupNode
) -> None:
    """A Const whose value, or a Dim whose bounds, name a variable: "Constant
    expression required". A ReDim is a statement and never reaches here."""
    for decl in group.declarations:
        if is_inactive_node(ctx.activity, decl):
            continue
        toks = statement_tokens(ctx.source, decl.span)
        read: list[VbaToken] = []
        if group.is_const:
            eq = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            read = [] if eq < 0 else toks[eq + 1 :]
        elif decl.fixed_length is not None:
            # `Dim s As String * L` (XLIDE issue #451, measured in Excel 16.0).
            star = next((k for k, tok in enumerate(toks) if tok.raw_text == "*"), -1)
            read = [] if star < 0 else toks[star + 1 :]
        elif decl.is_array and js_trim(decl.array_bounds or "") != "":
            open_ = next((k for k, tok in enumerate(toks) if tok.raw_text == "("), -1)
            depth = 0
            if open_ >= 0:
                for i in range(open_, len(toks)):
                    raw = toks[i].raw_text
                    depth += 1 if raw == "(" else -1 if raw == ")" else 0
                    if depth == 0:
                        read = toks[open_ + 1 : i]
                        break
        for i, tok in enumerate(read):
            name = token_name(tok)
            if (
                not name
                or _raw_at(read, i - 1) == "."
                or _raw_at(read, i + 1) == "."
                or _raw_at(read, i + 1) == "("
            ):
                continue
            if _variable_named(ctx, proc_sym, name) is None:
                continue
            if group.is_const:
                message = (
                    f"Const '{decl.name}' takes its value from the variable '{name}'. This is a "
                    "VBE compile error: Constant expression required."
                )
            elif decl.fixed_length is not None:
                message = (
                    f"The length of the fixed-length String '{decl.name}' names the variable "
                    f"'{name}'; it must be a constant. This is a VBE compile error: Constant "
                    "expression required."
                )
            else:
                message = (
                    f"The bounds of '{decl.name}' name the variable '{name}'; a Dim needs "
                    "constants there, and ReDim takes a variable. This is a VBE compile error: "
                    "Constant expression required."
                )
            ctx.push(
                "constValueNotConstant" if group.is_const else "arrayBoundNotConstant",
                message,
                absolute_span(decl.span, tok),
            )


# The type each type-declaration character stands for.
_SUFFIX_TYPES = {
    "%": "integer",
    "&": "long",
    "^": "longlong",
    "@": "currency",
    "!": "single",
    "#": "double",
    "$": "string",
}


def _check_type_suffixes(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    """`n% = 2` with n As Long: the character names another type than the declaration."""
    for i in range(len(toks) - 1):
        name_tok = toks[i]
        suffix_tok = toks[i + 1]
        suffix_type = _SUFFIX_TYPES.get(suffix_tok.raw_text)
        name = token_name(name_tok)
        if (
            not suffix_type
            or not name
            or suffix_tok.start != name_tok.end
            or _raw_at(toks, i - 1) == "."
        ):
            continue
        # `rs!Field` is a member, not a suffix.
        after = _at(toks, i + 2)
        if (
            after is not None
            and after.start == suffix_tok.end
            and (token_name(after) or after.kind is TokenKind.INTEGER_LITERAL)
        ):
            continue
        variable = _variable_named(ctx, proc_sym, name)
        if variable is None or variable.is_array:
            continue
        declared = normalize_type(variable.as_type)
        if declared is None and ctx.untyped_is_variant:
            declared = "variant"
        if not declared or declared == suffix_type:
            continue
        ctx.push(
            "typeSuffixMismatch",
            f"'{name}{suffix_tok.raw_text}' says {suffix_type}, but '{name}' is declared "
            f"{variable.as_type if variable.as_type is not None else 'Variant'}. This is a VBE "
            "compile error: Type-declaration character does not match declared data type.",
            Span(span.start + name_tok.start, span.start + suffix_tok.end),
        )


# The functions that take no named arguments: the ones VBA compiles as keywords.
_NO_NAMED_ARGUMENTS = frozenset(
    {
        "instr", "instrb", "len", "lenb", "strcomp", "abs", "int", "fix", "sgn",
        "cstr", "cint", "clng", "cdbl", "cbool", "cdate", "cvar", "cbyte", "ccur", "csng",
        "clnglng", "clngptr",
    }
)  # fmt: skip


def _check_named_arguments(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    for i in range(len(toks) - 1):
        name = token_name(toks[i])
        if (
            not name
            or name.lower() not in _NO_NAMED_ARGUMENTS
            or toks[i + 1].raw_text != "("
            or _raw_at(toks, i - 1) == "."
        ):
            continue
        # A project procedure of the same name takes named arguments like any other.
        binding = source_identifier_binding(
            ctx.symbols, proc_sym, ctx.project_visible_symbols, name, BareIdentifierContext.CALL
        )
        if binding.scope is not BareIdentifierResolutionScope.UNRESOLVED:
            continue
        depth = 0
        for j in range(i + 1, len(toks)):
            raw = toks[j].raw_text
            depth += 1 if raw == "(" else -1 if raw == ")" else 0
            if depth == 0:
                break
            if depth == 1 and raw == ":=":
                ctx.push(
                    "namedArgumentNotAllowed",
                    f"{name} takes its arguments by position only. This is a VBE compile error: "
                    "Syntax error.",
                    Span(span.start + toks[j - 1].start, span.start + toks[j].end),
                )
                break


def _check_mid_target(
    ctx: _Context, proc_sym: VbaSymbol | None, toks: Sequence[VbaToken], span: Span
) -> None:
    """`Mid$(S, 1, 1) = "x"` with S a Const."""
    head = token_text(_at(toks, 0))
    if head != "mid" and head != "midb":
        return
    open_ = 2 if _raw_at(toks, 1) == "$" else 1
    target = _at(toks, open_ + 1)
    name = token_name(target) if target is not None else None
    if _raw_at(toks, open_) != "(" or not name or _raw_at(toks, open_ + 2) != ",":
        return
    depth = 0
    for j in range(open_, len(toks)):
        raw = toks[j].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if depth == 0:
            if (
                _raw_at(toks, j + 1) == "="
                and target is not None
                and _is_constant_name(ctx, proc_sym, name)
            ):
                _variable_required(ctx, name, absolute_span(span, target))
            return
