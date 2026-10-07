"""Rule family: declarations the VBE refuses by their form (XLIDE issue #212).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/refusedDeclarations.ts.
Measured in 64-bit Excel 16.0 (build 20326, 2026-09-30), by compiling the
whole project:

- private-type-in-public-signature: a Private Enum of a standard module
  in a Public (or unmarked) Sub, Function, Property, Declare or variable;
  in a class module a Private Enum or Type in a Public procedure or Event,
  and a Private Enum as a Public variable. A Private Type is fine in a
  standard module's public signatures, and a Public Type's field may be of
  a Private Enum.
- optional-property-value: a Property Let or Set whose value parameter is
  Optional, "Syntax error". An Optional index before it compiles.
- event-parameter-form: an Event parameter that is Optional or a
  ParamArray, "Syntax error".
- const-invalid-type: a Const declared As a type that is not one of VBA's
  own: As Object is "Invalid data type for constant", and As Collection,
  Range, a class, an Enum, a Type or Decimal is "Expected: type name".
- empty-enum: an Enum with no members, "Enum without members not allowed".
- type-member-without-type: a Type member with no As clause, a type
  suffix included, "Statement invalid inside Type block".
- type-enum-name-conflict: a Type and an Enum of one name in one module,
  "Ambiguous name detected". A Const may share either's name.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from ...conditional import ConditionalActivityTracker
from ...js_compat import JS_WHITESPACE, js_trim
from ...parser.nodes import (
    DeclareNode,
    EnumNode,
    EventNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    TypeNode,
    VariableDeclNode,
    VariableGroupNode,
)
from ...symbols.symbol_model import ModuleSymbolKind
from ..context import PushFn
from ..walker import (
    active_module_members,
    declared_name_span,
    for_each_variable_group,
    is_inactive_node,
    match_paren_from,
    statement_tokens,
    token_text,
)

_PRIVATE_TYPE_MESSAGE = (
    "Private Enum and user defined types cannot be used as parameters or return types for "
    "public procedures, public data members, or fields of public user defined types."
)

# The types a Const may be declared As, and Decimal, which is refused here as
# well but is invalid-as-type-name's to report wherever it appears.
_CONST_TYPES = frozenset(
    {
        "boolean", "byte", "integer", "long", "longlong", "longptr", "currency", "single",
        "double", "date", "string", "variant", "decimal",
    }
)  # fmt: skip

_DIGITS_RE = re.compile(r"[0-9]+")
_EMPTY_PARENS_END_RE = re.compile(f"\\([{JS_WHITESPACE}]*\\)\\Z")
_PRIVATE_OR_FRIEND_RE = re.compile(r"(?:private|friend)", re.IGNORECASE | re.ASCII)
_PUBLIC_OR_GLOBAL_RE = re.compile(r"(?:public|global)", re.IGNORECASE | re.ASCII)


def check_refused_declarations(
    source: str,
    mod: ModuleNode,
    module_kind: ModuleSymbolKind,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    _check_private_types_in_public_signatures(source, mod, module_kind, activity, push)
    type_names: dict[str, Span] = {}
    enum_names: dict[str, Span] = {}

    # Two Enums or two Types of one name: Ambiguous name detected (XLIDE issue
    # #639, measured in Excel 16.0).
    def repeated_name(seen: dict[str, Span], name: str, span: Span, kinds: str) -> None:
        if name.lower() in seen:
            push(
                "typeEnumNameConflict",
                f"Two {kinds} in this module are both named '{name}'. This is a VBE compile "
                "error: Ambiguous name detected.",
                span,
            )
        else:
            seen[name.lower()] = span

    def const_groups(group: VariableGroupNode) -> None:
        if group.is_const:
            _check_const_types(source, group.declarations, activity, push)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            if (
                member.proc_kind in (ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET)
                and member.params
                and member.params[-1].optional
            ):
                value = member.params[-1]
                which = "Let" if member.proc_kind is ProcKind.PROPERTY_LET else "Set"
                push(
                    "optionalPropertyValue",
                    f"The value parameter '{value.name}' of a Property {which} cannot be "
                    "Optional. This is a VBE compile error.",
                    value.name_span if value.name_span is not None else value.span,
                )
            for_each_variable_group(member.body, const_groups, activity)
        elif isinstance(member, EventNode):
            for param in member.params:
                if param.optional or param.param_array:
                    what = "a ParamArray" if param.param_array else "Optional"
                    push(
                        "eventParameterForm",
                        f"An Event parameter cannot be {what}: '{param.name}' in '{member.name}'. "
                        "This is a VBE compile error: Syntax error.",
                        param.name_span if param.name_span is not None else param.span,
                    )
                elif param.is_array and param.by_val:
                    # An array passes ByRef (XLIDE issue #266, measured).
                    push(
                        "eventParameterForm",
                        f"An Event's array parameter must be ByRef: '{param.name}' in "
                        f"'{member.name}'. This is a VBE compile error: Array argument must be "
                        "ByRef.",
                        param.name_span if param.name_span is not None else param.span,
                    )
            # An Event is Public: `Private Event` and `Friend Event` are
            # "Expected: Sub or Function or Property", and an As clause after
            # it is "Expected: end of statement" (XLIDE issue #266, measured in a
            # class). In a standard module any Event is
            # event-declaration-module-kind's.
            if module_kind is ModuleSymbolKind.STANDARD:
                continue
            toks = statement_tokens(source, member.span)
            head = token_text(toks[0] if toks else None)
            if head == "private" or head == "friend":
                push(
                    "invalidProcedureHeader",
                    f"An Event cannot be {toks[0].raw_text}; an Event is always Public. This is a "
                    "VBE compile error: Expected: Sub or Function or Property.",
                    Span(member.span.start + toks[0].start, member.span.start + toks[0].end),
                )
            open_ = next((k for k, tok in enumerate(toks) if tok.raw_text == "("), -1)
            close = -1 if open_ < 0 else match_paren_from(toks, open_)
            after = None if close < 0 or close + 1 >= len(toks) else toks[close + 1]
            if after is not None and token_text(after) == "as":
                push(
                    "invalidProcedureHeader",
                    f"An Event returns nothing, so it takes no As clause: '{member.name}'. This is "
                    "a VBE compile error: Expected: end of statement.",
                    Span(member.span.start + after.start, member.span.start + toks[-1].end),
                )
        elif isinstance(member, VariableGroupNode):
            if member.is_const:
                _check_const_types(source, member.declarations, activity, push)
            # Static keeps a local's value between calls, and means nothing
            # outside a procedure (XLIDE issue #216).
            if member.modifier.lower() == "static":
                push(
                    "staticOutsideProcedure",
                    "Static declares a variable inside a procedure; at module level use Private "
                    "or Dim. This is a VBE compile error: Invalid outside procedure.",
                    Span(member.span.start, member.span.start + len("Static")),
                )
        elif isinstance(member, EnumNode):
            if member.closed and not any(
                not is_inactive_node(activity, enum_member) for enum_member in member.members
            ):
                push(
                    "emptyEnum",
                    f"Enum '{member.name}' must declare at least one member. This is a VBE "
                    "compile error: Enum without members not allowed.",
                    member.name_span if member.name_span is not None else member.span,
                )
            repeated_name(
                enum_names,
                member.name,
                member.name_span if member.name_span is not None else member.span,
                "Enums",
            )
        elif isinstance(member, TypeNode):
            for field in member.fields:
                if not field.has_as_clause and not is_inactive_node(activity, field):
                    # `10  Id As Long` reads as a member named 10: a line
                    # number, which a Type block does not take either.
                    push(
                        "typeMemberWithoutType",
                        "A line number cannot label a Type member. This is a VBE compile error: "
                        "Statement invalid inside Type block."
                        if _DIGITS_RE.fullmatch(field.name)
                        else f"Type member '{field.name}' needs an As clause. This is a VBE "
                        "compile error: Statement invalid inside Type block.",
                        field.name_span if field.name_span is not None else field.span,
                    )
            repeated_name(
                type_names,
                member.name,
                member.name_span if member.name_span is not None else member.span,
                "Types",
            )
    for name, span in enum_names.items():
        type_span = type_names.get(name)
        if type_span is not None:
            later = type_span if type_span.start > span.start else span
            push(
                "typeEnumNameConflict",
                f"A Type and an Enum in this module are both named "
                f"'{source[later.start : later.end]}'. This is a VBE compile error: Ambiguous "
                "name detected.",
                later,
            )


def _check_const_types(
    source: str,
    declarations: Sequence[VariableDeclNode],
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    for decl in declarations:
        as_type = js_trim(decl.as_type) if decl.as_type is not None else None
        if not as_type or is_inactive_node(activity, decl) or as_type.lower() in _CONST_TYPES:
            continue
        object_ = as_type.lower() == "object"
        push(
            "constInvalidType",
            "A Const cannot be declared As Object. This is a VBE compile error: Invalid data "
            "type for constant."
            if object_
            else f"A Const can only be declared As one of VBA's own types, not As {as_type}. "
            "This is a VBE compile error: Expected: type name.",
            declared_name_span(source, decl.span, decl.name),
        )


def _is_public_member(modifiers: Sequence[str]) -> bool:
    """Whether a procedure-like member is reachable from outside its module."""
    return not any(_PRIVATE_OR_FRIEND_RE.fullmatch(modifier) for modifier in modifiers)


def _check_private_types_in_public_signatures(
    source: str,
    mod: ModuleNode,
    module_kind: ModuleSymbolKind,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if module_kind is not ModuleSymbolKind.STANDARD and module_kind is not ModuleSymbolKind.CLASS:
        return
    private_enums: set[str] = set()
    private_types: set[str] = set()
    for member in active_module_members(mod, activity):
        if (
            isinstance(member, EnumNode)
            and member.visibility is not None
            and member.visibility.lower() == "private"
        ):
            private_enums.add(member.name.lower())
        elif (
            isinstance(member, TypeNode)
            and member.visibility is not None
            and member.visibility.lower() == "private"
        ):
            private_types.add(member.name.lower())

    # A standard module may use its Private Types publicly; a class may not.
    def refused(type_name: str | None) -> bool:
        if type_name is None:
            return False
        key = js_trim(_EMPTY_PARENS_END_RE.sub("", type_name, count=1)).lower()
        return key in private_enums or (
            module_kind is ModuleSymbolKind.CLASS and key in private_types
        )

    if len(private_enums) == 0 and (
        module_kind is not ModuleSymbolKind.CLASS or len(private_types) == 0
    ):
        return

    def report(span: Span) -> None:
        push(
            "privateTypeInPublicSignature",
            f"{_PRIVATE_TYPE_MESSAGE} This is a VBE compile error.",
            span,
        )

    def check_params(params: Sequence[ParameterNode]) -> None:
        for param in params:
            if refused(param.as_type):
                report(param.name_span if param.name_span is not None else param.span)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode) and _is_public_member(member.modifiers):
            check_params(member.params)
            if refused(member.return_type):
                report(member.name_span if member.name_span is not None else member.span)
        elif isinstance(member, DeclareNode) and _is_public_member(
            [member.visibility if member.visibility is not None else ""]
        ):
            check_params(member.params)
            if refused(member.return_type):
                report(member.name_span if member.name_span is not None else member.span)
        elif isinstance(member, EventNode) and _is_public_member(
            [member.visibility if member.visibility is not None else ""]
        ):
            check_params(member.params)
        elif (
            isinstance(member, VariableGroupNode)
            and not member.is_const
            and _PUBLIC_OR_GLOBAL_RE.fullmatch(member.modifier)
        ):
            for decl in member.declarations:
                # A class's Public variable of its Private Type is the object-module
                # rule's; only the Enum is this message there.
                key = js_trim(decl.as_type).lower() if decl.as_type is not None else None
                if (
                    key is not None
                    and key in private_enums
                    and not is_inactive_node(activity, decl)
                ):
                    report(declared_name_span(source, decl.span, decl.name))
