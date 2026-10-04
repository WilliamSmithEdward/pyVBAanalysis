"""Rule family: the members of a user-defined type (XLIDE issue #253).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/typeMembers.ts. Every
case was measured in Excel 16.0 (build 20326, 2026-10-01).

 - object-variable-not-set: an object member of a Type local that nothing has
   Set, or that was Set to Nothing, given a member or a call: `t.o.Add 1`,
   `t.o.Count`, `t.o(1)`, `.o.Add 1` inside `With t`, and `.Add 1` inside
   `With t.o`. Each raises 91.
 - scalar-member-access: a member of a member that holds a number or a string,
   `t.a.Value`, `t.s.Length`, `.a.Value` in `With t`: "Invalid qualifier".
 - fixed-array-redim: ReDim of a fixed array member, `ReDim t.f(3)`: "Array
   already dimensioned".
 - erase-requires-array: Erase of a whole Type value, `Erase t`, or of a member
   that is no array, `Erase t.a`: "Expected array".

What a member holds comes from type_member_state.py.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence

from ...completion.member_access import MemberCompletionContext, is_known_object_assignment_type
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, VbaSymbolKind
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn, statement_tokens
from ..type_fields import (
    FieldStep,
    ModuleTypes,
    WithSubject,
    field_chain,
    is_fixed_array_field,
    is_leading_dot,
    module_types,
    type_key,
    type_root_at,
    variable_symbol_in,
    walk_with_subjects,
)
from ..type_member_state import MemberStatesAt, type_member_states_at
from ..walker import active_module_members, statement_tokens_after_leading_label, token_name, token_text
from .arrays import module_option_base

_NOT_SET = "This will raise Run-time error '91': Object variable or With block variable not set."

# (rule, message, span)
_Hit = tuple[str, str, Span]


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _js(value: str | None) -> str:
    """A string or undefined, as a JavaScript template literal prints it."""
    return "undefined" if value is None else value


def check_type_members(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    types = module_types(source, mod, activity)
    if not types:
        return
    option_base = module_option_base(mod, activity)

    def is_object_type(type_: str) -> bool:
        return type_ not in types and type_ != "variant" and is_known_object_assignment_type(type_, member_ctx)

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        states_at = type_member_states_at(source, symbols, member, types, activity, option_base, is_object_type)
        _check_for_each_fields(source, member.body, symbols, member, types, is_object_type, states_at, activity, push)

        def visit(
            stmt: LeafStatementNode,
            subject: WithSubject | None,
            member: ProcedureNode = member,
            states_at: MemberStatesAt = states_at,
        ) -> None:
            toks = statement_tokens_after_leading_label(source, stmt.span)
            head = token_text(_at(toks, 0))
            if head in ("redim", "erase"):
                _check_resizes(stmt.span, toks, head, symbols, member, types, subject, is_object_type, push)
                if head == "erase":
                    _check_erase_of_empty(stmt, toks, symbols, member, types, subject, states_at, push)
                return
            _check_with_object(stmt, toks, subject, states_at, push)
            for i in range(len(toks)):
                root = type_root_at(toks, i, symbols, member, types, subject)
                if root is None:
                    continue
                steps = field_chain(toks, root, types)
                hit = (
                    _scalar_qualifier(stmt.span, toks, i, steps)
                    or _nothing_access(stmt, toks, i, steps, states_at, is_object_type)
                    or _field_use_misuse(stmt.span, toks, i, steps, types, is_object_type, symbols, member)
                    or _held_value_misuse(stmt, toks, i, steps, states_at)
                )
                if hit is not None:
                    push(hit[0], hit[1], hit[2])

        walk_with_subjects(source, member.body, activity, symbols, member, types, None, visit)


def _field_kind(step: FieldStep, types: ModuleTypes, is_object_type: Callable[[str], bool]) -> str | None:
    """What the last field of a chain gives: 'array' (a whole array), 'scalar'
    (a value of a scalar type), 'udt' (a Type), 'object', or 'variant'."""
    if step.field.is_array and step.open is None:
        return "array"
    type_ = step.field.type
    normal = normalize_type(type_)
    if not type_ or normal == "variant":
        return "variant"
    if type_ in types:
        return "udt"
    if is_known_scalar_type(normal or ""):
        return "scalar"
    return "object" if is_object_type(type_) else None


_VALUE_OPERATORS = frozenset({"+", "-", "*", "/", "\\", "^", "mod", "&", "=", "<>", "<", ">", "<=", ">="})
_EMPTY_PARENS_END_RE = re.compile(r"\(\)\Z")
_NUMBER_KINDS = frozenset({TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL})
_LITERAL_KINDS = frozenset({TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL})


def _field_use_misuse(
    span: Span,
    toks: Sequence[VbaToken],
    start: int,
    steps: Sequence[FieldStep],
    types: ModuleTypes,
    is_object_type: Callable[[str], bool],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
) -> _Hit | None:
    """A field of a Type used as the shape it is not (issue #417, each measured
    in Excel 16.0): what the same misuse of a local reports, for `t.f`.

     - `t.f(1)` on a number, a string or a Type: "Expected array"
       (scalar-indexed). `UBound(t.f)` on what is no array: "Expected array".
     - `t.f Is Nothing` on a value, a Type or an array: "Type mismatch".
     - `t.f = 5` on an array: "Can't assign to array"; on a Type: "Type
       mismatch". `Set t.f = ...` on a Type: "Object required".
     - `t.f + 1`, `t.f & "x"` on an array or a Type: "Type mismatch"; on a
       Collection, and `Len(t.f)`: "Argument not optional".
     - `t.f.Count` on an array: "Invalid qualifier".
     - A Type read whole into a Variant, `Main = t.f` or a ByVal Variant
       argument: "Only user-defined types defined in public object modules can
       be coerced to or from a variant". A Type, an array or a Variant passed to
       a ByRef parameter of another type: "ByRef argument type mismatch".
    """
    step = steps[-1] if steps else None
    if step is None:
        return None
    end = step.close if step.close is not None else step.at
    at = Span(span.start + toks[start].start, span.start + toks[end].end)
    shown = step.display
    declared = (
        step.field.type_name
        if step.field.type_name is not None
        else step.field.type
        if step.field.type is not None
        else "Variant"
    )
    after = _at(toks, end + 1)
    before = _at(toks, start - 1)
    word = token_text
    # field_chain stops at the parenthesis after a field that is no array.
    if not step.field.is_array and _raw_at(toks, step.at + 1) == "(":
        type_ = step.field.type
        if type_ and (
            type_ in types or (normalize_type(type_) != "variant" and is_known_scalar_type(normalize_type(type_) or ""))
        ):
            return (
                "scalarIndexed",
                f"'{shown}' is declared As {declared}, which is no array to index. "
                "This is a VBE compile error: Expected array.",
                at,
            )
        return None
    kind = _field_kind(step, types, is_object_type)
    if not kind or (after is not None and after.raw_text == "." and kind != "array"):
        return None
    whole_argument = after is None or after.raw_text in (")", ",")
    if (
        before is not None
        and before.raw_text == "("
        and word(_at(toks, start - 2)) in ("ubound", "lbound")
        and whole_argument
        and kind != "array"
        and kind != "variant"
    ):
        return (
            "arrayBoundRequiresArray",
            f"{toks[start - 2].raw_text} takes an array, and '{shown}' is declared As {declared}. "
            "This is a VBE compile error: Expected array.",
            at,
        )
    if (word(after) == "is" or (word(before) == "is" and word(_at(toks, start - 2)) != "typeof")) and kind in (
        "scalar",
        "udt",
        "array",
    ):
        what = "an array" if kind == "array" else f"declared As {declared}"
        return (
            "isOperatorNonObject",
            f"The 'Is' operator requires object operands, but '{shown}' is {what}. "
            "This is a VBE compile error: Type mismatch.",
            at,
        )
    head = word(_at(toks, 0))
    target = (start == 0 or (start == 1 and head == "let")) and after is not None and after.raw_text == "="
    value = list(toks[end + 2 :]) if target else []
    # A dynamic array field takes an array or, As Byte, a String: only a number
    # literal is judged there (issue #417, measured in Excel 16.0).
    number_value = len(value) == 1 and value[0].kind in _NUMBER_KINDS
    # A String into one that is not As Byte is refused too (issue #604).
    element_type = normalize_type(step.field.type)
    string_value = len(value) == 1 and value[0].kind is TokenKind.STRING_LITERAL and element_type != "byte"
    if target and kind == "array" and (is_fixed_array_field(step.field) or number_value or string_value):
        return (
            "arrayTargetAssignment",
            f"'{shown}' is an array, which a value cannot be assigned to whole. "
            "This is a VBE compile error: Can't assign to array.",
            at,
        )
    # `t.f = Split("a b")` gives a String array to a number array: 13 (issue
    # #604, measured in Excel 16.0).
    if (
        target
        and kind == "array"
        and element_type is not None
        and is_known_scalar_type(element_type)
        and element_type not in ("string", "byte", "boolean", "date")
        and word(_at(value, 0)) == "split"
        and _raw_at(value, 1) == "("
        and match_paren_from(value, 1) == len(value) - 1
    ):
        return (
            "assignmentTypeMismatch",
            f"'{shown}' is an array of {_EMPTY_PARENS_END_RE.sub('', declared, count=1)}, and Split gives an "
            "array of String. This will raise Run-time error '13': Type mismatch.",
            at,
        )
    if target and kind == "udt" and len(value) == 1 and value[0].kind in _LITERAL_KINDS:
        return (
            "udtValueMismatch",
            f"'{shown}' is a {declared}, which {value[0].raw_text} cannot be assigned to. "
            "This is a VBE compile error: Type mismatch.",
            at,
        )
    if start == 1 and head == "set" and after is not None and after.raw_text == "=" and kind == "array":
        return (
            "arrayTargetAssignment",
            f"'{shown}' is an array, which Set cannot assign to. This is a VBE compile error: Can't assign to array.",
            at,
        )
    if start == 1 and head == "set" and after is not None and after.raw_text == "=" and kind == "udt":
        return (
            "setRequiresObject",
            f"'{shown}' is a {declared}, no object variable for Set. This is a VBE compile error: Object required.",
            at,
        )
    assign_at = next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
    operator_before = before is not None and word(before) in _VALUE_OPERATORS and start - 1 != assign_at
    operator_after = after is not None and word(after) in _VALUE_OPERATORS and not target and end + 1 != assign_at
    if operator_before or operator_after:
        operator = after if operator_after else before
        assert operator is not None
        if kind in ("array", "udt"):
            what = "an array" if kind == "array" else f"a {declared}"
            return (
                "nonScalarBinaryOperand",
                f"The '{operator.raw_text}' operator requires a scalar operand, but '{shown}' is {what}. "
                "This is a VBE compile error: Type mismatch.",
                at,
            )
        if kind == "object" and normalize_type(step.field.type) == "collection":
            return (
                "collectionOperand",
                f"'{shown}' is a Collection: its default member Item needs an index, so '{operator.raw_text}' has "
                "no value to work on. This is a VBE compile error: Argument not optional.",
                at,
            )
    if (
        kind == "object"
        and normalize_type(step.field.type) == "collection"
        and before is not None
        and before.raw_text == "("
        and word(_at(toks, start - 2)) == "len"
        and after is not None
        and after.raw_text == ")"
    ):
        return (
            "collectionOperand",
            f"'{shown}' is a Collection: its default member Item needs an index, so Len has no value to take. "
            "This is a VBE compile error: Argument not optional.",
            at,
        )
    if kind == "array" and after is not None and after.raw_text == "." and token_name(_at(toks, end + 2)):
        return (
            "scalarMemberAccess",
            f"Member access on '{shown}' is invalid because it is an array. "
            "This is a VBE compile error: Invalid qualifier.",
            Span(at.start, span.start + toks[end + 2].end),
        )
    # A Type read whole into a Variant: `Main = t.f`.
    if kind == "udt" and assign_at == 1 and start == 2 and after is None:
        lhs_name = token_name(_at(toks, 0))
        lhs = lhs_name.lower() if lhs_name else None
        local = variable_symbol_in(symbols, proc, lhs) if lhs else None
        if lhs == proc.name.lower():
            into_variant = not proc.return_type or normalize_type(proc.return_type) == "variant"
        else:
            into_variant = (
                local is not None
                and not local.is_array
                and (not local.as_type or normalize_type(local.as_type) == "variant")
            )
        if into_variant:
            return (
                "udtVariantCoercion",
                f"'{shown}' is a {declared}, which a Variant cannot hold. This is a VBE compile error: Only "
                "user-defined types defined in public object modules can be coerced to or from a variant or "
                "passed to late-bound functions.",
                at,
            )
    # An argument of a call statement to a procedure of the module: `TakeV t.f`.
    callee_name = token_name(_at(toks, 0))
    callee = callee_name.lower() if callee_name else None
    procedure = (
        next(
            (
                child
                for child in (symbols.root.children or [])
                if child.kind in (VbaSymbolKind.SUB, VbaSymbolKind.FUNCTION) and child.name.lower() == callee
            ),
            None,
        )
        if callee
        else None
    )
    if (
        procedure is not None
        and start >= 1
        and (
            (before is not None and before.raw_text == ",")
            or start == 1
            or (start == 2 and before is not None and before.raw_text == "(")
        )
        and whole_argument
    ):
        position = sum(1 for tok in toks[1:start] if tok.raw_text == ",")
        parameters = [child for child in (procedure.children or []) if child.kind is VbaSymbolKind.PARAMETER]
        param = parameters[position] if position < len(parameters) else None
        param_type = normalize_type(param.as_type if param is not None else None)
        if param is not None and not param.param_array and not param.is_array:
            if kind == "udt" and param.by_val and (not param_type or param_type == "variant"):
                return (
                    "udtVariantCoercion",
                    f"'{shown}' is a {declared}, which the ByVal Variant '{param.name}' of '{procedure.name}' "
                    "cannot take. This is a VBE compile error: Only user-defined types defined in public object "
                    "modules can be coerced to or from a variant or passed to late-bound functions.",
                    at,
                )
            if (
                not param.by_val
                and param_type
                and param_type != "variant"
                and (
                    kind == "array"
                    or (kind in ("udt", "variant") and param_type != normalize_type(step.field.type))
                )
            ):
                what = "an array" if kind == "array" else f"declared As {declared}"
                return (
                    "byRefArgumentTypeMismatch",
                    f"ByRef argument '{param.name}' of '{procedure.name}' expects {_js(param.as_type)}, but "
                    f"'{shown}' is {what}. This is a VBE compile error: ByRef argument type mismatch.",
                    at,
                )
    return None


_ARITHMETIC = frozenset({"+", "-", "*", "/", "\\", "^", "mod"})


def _held_value_misuse(
    stmt: LeafStatementNode,
    toks: Sequence[VbaToken],
    start: int,
    steps: Sequence[FieldStep],
    states_at: MemberStatesAt,
) -> _Hit | None:
    """A field read for what it still holds from the Dim (issue #417, each
    measured in Excel 16.0): an object field nothing has Set, read as a value,
    raises 91; a Collection Set to a New Collection raises 450 read whole and 5
    given an index before anything is added; a String field still "" next to a
    number in arithmetic raises 13; a Variant field still Empty raises 13 given
    an index and 424 given a member."""
    step = steps[-1] if steps else None
    if step is None or not step.path or step.open is not None:
        return None
    state = states_at(stmt, stmt.span.start + toks[start].start).get(step.path)
    if state not in ("nothing", "emptyCollection", "emptyString", "empty"):
        return None
    end = step.at
    at = Span(stmt.span.start + toks[start].start, stmt.span.start + toks[end].end)
    shown = step.display
    head = token_text(_at(toks, 0))
    before = _at(toks, start - 1)
    after = _at(toks, end + 1)
    # The `=` of an assignment statement, not of a comparison.
    first = _at(toks, 0)
    assign_at = (
        next((i for i, tok in enumerate(toks) if tok.raw_text == "="), -1)
        if (first is not None and first.kind is TokenKind.IDENTIFIER) or head in ("let", "set")
        else -1
    )
    if start + (-1 if head in ("let", "set") else 0) == 0 and after is not None and after.raw_text == "=":
        return None
    operator_before = before is not None and token_text(before) in _VALUE_OPERATORS and start - 1 != assign_at
    operator_after = after is not None and token_text(after) in _VALUE_OPERATORS and end + 1 != assign_at
    let_read = head != "set" and assign_at > 0 and start == assign_at + 1 and after is None
    call = (
        token_text(_at(toks, start - 2))
        if before is not None and before.raw_text == "(" and after is not None and after.raw_text == ")"
        else ""
    )
    is_operand = token_text(after) == "is" or (
        token_text(before) == "is" and token_text(_at(toks, start - 2)) != "typeof"
    )
    if state == "nothing" and (operator_before or operator_after or let_read or call in ("len", "lenb")):
        return (
            "objectVariableNotSet",
            f"'{shown}' is an object member that nothing has Set here, so it has no value to read. {_NOT_SET}",
            at,
        )
    if state == "empty" and call in ("ubound", "lbound"):
        return (
            "variantValueMisuse",
            f"'{shown}' is a Variant that still holds Empty, which is no array for {toks[start - 2].raw_text}. "
            "This will raise Run-time error '13': Type mismatch.",
            at,
        )
    if state == "empty" and is_operand:
        return (
            "variantValueMisuse",
            f"'{shown}' is a Variant that still holds Empty, not an object, so Is cannot compare it. "
            "This will raise Run-time error '424': Object required.",
            at,
        )
    if state == "emptyCollection" and let_read:
        return (
            "objectDefaultValue",
            f"'{shown}' is a Collection: its default member Item needs an index, so it has no value to read "
            "here. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.",
            at,
        )
    if state == "emptyCollection" and after is not None and after.raw_text == "(":
        return (
            "collectionIndexOutOfRange",
            f"'{shown}' holds nothing here, so no index reaches an element. "
            "This will raise Run-time error '5': Invalid procedure call or argument.",
            at,
        )
    if operator_after and token_text(after) in _ARITHMETIC:
        number_at = end + 2
    elif operator_before and token_text(before) in _ARITHMETIC:
        number_at = start - 2
    else:
        number_at = -1
    number = _at(toks, number_at)
    if state == "emptyString" and number is not None and number.kind in _NUMBER_KINDS:
        held = 'holds ""' if step.field.fixed_length is None else "holds only spaces"
        operator = after if number_at > end else before
        assert operator is not None
        return (
            "stringArithmeticCoercion",
            f"Operator '{operator.raw_text}' coerces '{shown}', which {held}, to a number. "
            "This will raise Run-time error '13': Type mismatch.",
            at,
        )
    if state == "empty" and after is not None and after.raw_text == "(":
        return (
            "variantValueMisuse",
            f"'{shown}' is a Variant that still holds Empty, which takes no index. "
            "This will raise Run-time error '13': Type mismatch.",
            at,
        )
    if state == "empty" and after is not None and after.raw_text == "." and token_name(_at(toks, end + 2)):
        return (
            "variantValueMisuse",
            f"'{shown}' is a Variant that still holds Empty, not an object, so it has no "
            f"{toks[end + 2].raw_text}. This will raise Run-time error '424': Object required.",
            at,
        )
    return None


def _check_erase_of_empty(
    stmt: LeafStatementNode,
    toks: Sequence[VbaToken],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
    states_at: MemberStatesAt,
    push: PushFn,
) -> None:
    """`Erase t.v` while the Variant field still holds Empty: 13 (issue #417, measured in Excel 16.0)."""
    for i in range(1, len(toks)):
        if toks[i - 1].raw_text != "," and i != 1:
            continue
        root = type_root_at(toks, i, symbols, proc, types, subject)
        steps = field_chain(toks, root, types) if root is not None else []
        step = steps[-1] if steps else None
        after = _at(toks, (step.close if step.close is not None else step.at) + 1) if step is not None else None
        if (
            step is None
            or not step.path
            or step.open is not None
            or (after is not None and after.raw_text != "," and after.kind is not TokenKind.COMMENT)
        ):
            continue
        if states_at(stmt, stmt.span.start + toks[i].start).get(step.path) == "empty":
            push(
                "variantValueMisuse",
                f"'{step.display}' is a Variant that still holds Empty, which is not an array for Erase to act "
                "on. This will raise Run-time error '13': Type mismatch.",
                Span(stmt.span.start + toks[i].start, stmt.span.start + toks[step.at].end),
            )


def _check_for_each_fields(
    source: str,
    body: Sequence[BodyNode],
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    is_object_type: Callable[[str], bool],
    states_at: MemberStatesAt,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`For Each e In t.f` with f a number, a string or a Type (issue #417, measured in Excel 16.0)."""
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if not isinstance(node, ForBlockNode) or not node.each or node.source_expression_span is None:
            continue
        span = node.source_expression_span
        toks = statement_tokens(source, span)
        if not toks:
            continue
        root = type_root_at(toks, 0, symbols, proc, types, None)
        steps = field_chain(toks, root, types) if root is not None else []
        step = steps[-1] if steps else None
        kind = (
            _field_kind(step, types, is_object_type)
            if step is not None and (step.close if step.close is not None else step.at) == len(toks) - 1
            else None
        )
        state = states_at(node, span.start).get(step.path) if step is not None and step.path and kind else None
        if step is not None and kind in ("scalar", "udt"):
            push(
                "forEachSourceType",
                f"For Each source '{step.display}' must be a collection object or array, but it is declared As "
                f"{_js(step.field.type_name)}. This is a VBE compile error: For Each may only iterate over a "
                "collection object or an array.",
                span,
            )
        elif step is not None and state == "nothing":
            push(
                "objectVariableNotSet",
                f"'{step.display}' is an object member that nothing has Set here, so it is Nothing when For Each "
                "asks it for its elements. This will raise Run-time error '424': Object required.",
                span,
            )
        elif step is not None and state == "unallocated":
            push(
                "unallocatedDynamicArrayAccess",
                f"Dynamic array '{step.display}' is not allocated when For Each asks it for its elements. This "
                "will raise Run-time error '92': For loop not initialized.",
                span,
            )
        elif step is not None and state == "empty":
            push(
                "variantValueMisuse",
                f"'{step.display}' is a Variant that still holds Empty, which For Each cannot step through. This "
                "will raise Run-time error '13': Type mismatch.",
                span,
            )


def _scalar_qualifier(span: Span, toks: Sequence[VbaToken], start: int, steps: Sequence[FieldStep]) -> _Hit | None:
    """`t.a.Value`: a member of a field that holds a number or a string."""
    for step in steps:
        end = step.close if step.close is not None else step.at
        value = not step.field.is_array or step.open is not None
        type_ = normalize_type(step.field.type)
        if value and type_ and is_known_scalar_type(type_) and _raw_at(toks, end + 1) == ".":
            return (
                "scalarMemberAccess",
                f"Member access on '{step.display}' is invalid because it is declared as "
                f"{_js(step.field.type_name)}. This is a VBE compile error: Invalid qualifier.",
                Span(span.start + toks[start].start, span.start + toks[end + 1].end),
            )
    return None


def _nothing_access(
    stmt: LeafStatementNode,
    toks: Sequence[VbaToken],
    start: int,
    steps: Sequence[FieldStep],
    states_at: MemberStatesAt,
    is_object_type: Callable[[str], bool],
) -> _Hit | None:
    """`t.o.Add 1` and `t.o(1)` with t.o still Nothing."""
    for step in steps:
        nxt = _raw_at(toks, step.at + 1)
        if (
            step.field.is_array
            or not step.path
            or not is_object_type(step.field.type or "")
            or nxt not in (".", "!", "(")
        ):
            continue
        if states_at(stmt, stmt.span.start + toks[start].start).get(step.path) == "nothing":
            return (
                "objectVariableNotSet",
                f"'{step.display}' is an object member that nothing has Set here, so it is Nothing. {_NOT_SET}",
                Span(stmt.span.start + toks[start].start, stmt.span.start + toks[step.at].end),
            )
        return None
    return None


def _check_with_object(
    stmt: LeafStatementNode,
    toks: Sequence[VbaToken],
    subject: WithSubject | None,
    states_at: MemberStatesAt,
    push: PushFn,
) -> None:
    """`.Add 1` inside `With t.o`, while t.o is Nothing: reported once, at the first member it reaches."""
    if subject is None or not subject.path or subject.type or subject.field is None or subject.field.is_array:
        return
    for i in range(len(toks) - 1):
        if toks[i].raw_text != "." or not is_leading_dot(toks, i) or not token_name(toks[i + 1]):
            continue
        if states_at(stmt, stmt.span.start + toks[i].start).get(subject.path) == "nothing":
            push(
                "objectVariableNotSet",
                f"The With object '{subject.display}' is an object member that nothing has Set, so it is Nothing "
                f"here. {_NOT_SET}",
                Span(stmt.span.start + toks[i].start, stmt.span.start + toks[i + 1].end),
            )
        return


def _check_resizes(
    span: Span,
    toks: Sequence[VbaToken],
    head: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    types: ModuleTypes,
    subject: WithSubject | None,
    is_object_type: Callable[[str], bool],
    push: PushFn,
) -> None:
    """`ReDim t.f(3)` of a fixed member, and `Erase t` or `Erase t.a` of what is no array."""
    first = 2 if head == "redim" and token_text(_at(toks, 1)) == "preserve" else 1
    start = first
    for i in range(first, len(toks) + 1):
        if i < len(toks) and (toks[i].raw_text != "," or _depth_at(toks, start, i) > 0):
            continue
        group = [tok for tok in toks[start:i] if tok.kind is not TokenKind.COMMENT]
        start = i + 1
        if not group:
            continue

        def at(last: VbaToken, group: list[VbaToken] = group) -> Span:
            return Span(span.start + group[0].start, span.start + last.end)

        if head == "erase" and len(group) == 1:
            name = token_name(group[0])
            variable = variable_symbol_in(symbols, proc, name.lower() if name else "")
            type_ = type_key(variable.as_type) if variable is not None and not variable.is_array else None
            if type_ and type_ in types:
                push(
                    "eraseRequiresArray",
                    f"Erase target '{group[0].raw_text}' is a user-defined type, not an array. "
                    "This is a VBE compile error: Expected array.",
                    at(group[0]),
                )
            continue
        root = type_root_at(group, 0, symbols, proc, types, subject)
        steps = field_chain(group, root, types) if root is not None else []
        step = steps[-1] if steps else None
        if step is None:
            continue
        # What is no array and no Variant: a number, a string, a Type, a
        # Collection or an Object (issue #417, measured in Excel 16.0).
        field_type = step.field.type
        not_array = (
            not step.field.is_array
            and field_type is not None
            and normalize_type(field_type) != "variant"
            and (
                is_known_scalar_type(normalize_type(field_type) or "")
                or field_type in types
                or is_object_type(field_type)
            )
        )
        if head == "redim" and is_fixed_array_field(step.field) and step.open is not None:
            push(
                "fixedArrayRedim",
                f"'{step.display}' is a fixed-size array member, which ReDim cannot resize. "
                "This is a VBE compile error: Array already dimensioned.",
                at(group[step.at]),
            )
        elif head == "redim" and not_array and _raw_at(group, step.at + 1) == "(":
            push(
                "scalarRedim",
                f"ReDim target '{step.display}' must be an array or Variant, but it is declared As "
                f"{_js(step.field.type_name)}. This is a VBE compile error: Expected array.",
                at(group[step.at]),
            )
        elif head == "erase" and not_array and step.at == len(group) - 1:
            push(
                "eraseRequiresArray",
                f"Erase target '{step.display}' must be an array or Variant, but it is declared As "
                f"{_js(step.field.type_name)}. This is a VBE compile error: Expected array.",
                at(group[step.at]),
            )


def _depth_at(toks: Sequence[VbaToken], start: int, index: int) -> int:
    """The parenthesis depth at `index`, counting from `start`."""
    depth = 0
    for i in range(start, index):
        raw = toks[i].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
    return depth


