"""Rule family: declaration-site rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/declarations.ts. This
slice is the self-contained subset (procedure headers, identifier spelling,
reserved names, Dim initializers, unexpected declaration tokens, type-declaration
characters, Option placement/duplication, empty Type, parameter/identifier
limits, UDT parameter constraints). The rules that need type inference, the
member-completion context, constant-expression evaluation, or host/runtime
resolution (parameter defaults, fixed-length-string bounds, As-type-name
validation, property setter/accessor value types, non-constant values, parameter
order) are deferred to later slices / M8 / M9.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Union

from ...conditional import ConditionalActivity, ConditionalActivityTracker
from ...lexer.keyword_table import is_reserved_identifier
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.fixed_length_string import parse_fixed_length_string_type
from ...parser.nodes import (
    AttributeNode,
    ConditionalDirectiveNode,
    DeclareNode,
    EnumNode,
    ModuleNode,
    OptionNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    TypeFieldNode,
    TypeNode,
    VariableDeclNode,
    VariableGroupNode,
)
from ...parser.type_declaration_suffix import is_type_declaration_suffix
from ..context import PushFn, statement_tokens
from ..walker import (
    absolute_span,
    active_module_members,
    first_token_span,
    for_each_variable_group,
    is_inactive_node,
    strip_header_brackets,
    token_name,
    token_text,
)
from .shared import NameTokenHit, declaration_name_hit, name_token_hit

# Access/storage modifiers that may lead a procedure declaration.
_PROC_MODIFIERS: frozenset[str] = frozenset({"public", "private", "friend", "global", "static"})
_MAX_PROCEDURE_PARAMETERS = 60
_MAX_IDENTIFIER_LENGTH = 255

# Nodes that may carry a legacy type-declaration suffix plus an As clause.
_TypeDeclarationSuffixNode = Union[ParameterNode, ProcedureNode, TypeFieldNode, VariableDeclNode]


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _is_digit_started_token(tok: VbaToken) -> bool:
    return (tok.kind is TokenKind.INTEGER_LITERAL or tok.kind is TokenKind.FLOAT_LITERAL) and (
        len(tok.raw_text) > 0 and tok.raw_text[0].isdigit()
    )


def _first_line_span(source: str, span: Span) -> Span:
    nl = source.find("\n", span.start)
    return Span(span.start, span.end if nl == -1 else min(nl, span.end))


# -- checkProcedureHeader --------------------------------------------------


def check_procedure_header(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        header_start = member.span.start
        nl = source.find("\n", header_start)
        header_end = member.span.end if nl == -1 else nl
        toks = [
            t
            for t in tokenize(source[header_start:header_end])
            if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
        ]
        i = 0
        while i < len(toks) and toks[i].raw_text.lower() in _PROC_MODIFIERS:
            i += 1
        kw_tok = _at(toks, i)
        kw = kw_tok.raw_text.lower() if kw_tok is not None else None
        allow_as = False
        if kw == "function":
            allow_as = True
            i += 1
        elif kw == "sub":
            i += 1
        elif kw == "property":
            i += 1
            accessor = _at(toks, i)
            if accessor is not None and accessor.raw_text.lower() == "get":
                allow_as = True
            i += 1  # skip the accessor (Get/Let/Set)
        else:
            continue
        name_tok = _at(toks, i)
        if name_tok is None:
            continue
        if _is_digit_started_token(name_tok):
            continue  # invalid-identifier-start owns this range
        next_index = i + 1
        suffix_tok = _at(toks, next_index)
        if (
            allow_as
            and suffix_tok is not None
            and name_tok.end == suffix_tok.start
            and is_type_declaration_suffix(suffix_tok.raw_text)
        ):
            next_index += 1
        nxt = _at(toks, next_index)
        if nxt is None:
            continue
        r = nxt.raw_text
        if r == "(" or (allow_as and r.lower() == "as"):
            continue
        push(
            "invalidProcedureHeader",
            f"Unexpected '{r}' after procedure name '{strip_header_brackets(name_tok.raw_text)}'; "
            "a procedure name must be a single identifier.",
            Span(header_start + nxt.start, header_start + nxt.end),
        )


# -- checkInvalidIdentifierStarts ------------------------------------------


@dataclass(frozen=True, slots=True)
class _InvalidIdentifierStartHit:
    name: str
    span: Span
    reason: str  # "digit" | "underscore" | "hyphen" | "dot"


def _is_invalid_identifier_text_char(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch == "_")


def _invalid_identifier_text_end(source: str, start: int, limit: int) -> int:
    end = start
    while end < limit and _is_invalid_identifier_text_char(source[end]):
        end += 1
    return end


def _is_parameter_modifier(tok: VbaToken | None) -> bool:
    return token_text(tok) in ("optional", "byval", "byref", "paramarray")


def _invalid_identifier_start_at(
    source: str, base: Span, toks: Sequence[VbaToken], index: int
) -> _InvalidIdentifierStartHit | None:
    tok = _at(toks, index)
    if tok is None or tok.kind is TokenKind.BRACKETED_IDENTIFIER:
        return None
    # Embedded invalid character: an identifier directly followed by '-' or '.'.
    nxt = _at(toks, index + 1)
    if tok.kind is TokenKind.IDENTIFIER and nxt is not None and (nxt.raw_text == "-" or nxt.raw_text == "."):
        start = base.start + tok.start
        after = _at(toks, index + 2)
        end = base.start + (after.end if after is not None else nxt.end)
        return _InvalidIdentifierStartHit(
            name=source[start:end], span=Span(start, end), reason="hyphen" if nxt.raw_text == "-" else "dot"
        )
    reason: str | None = None
    if _is_digit_started_token(tok):
        reason = "digit"
    elif tok.raw_text.startswith("_"):
        reason = "underscore"
    if reason is None:
        return None
    start = base.start + tok.start
    end = _invalid_identifier_text_end(source, start, base.end)
    return _InvalidIdentifierStartHit(name=source[start:end], span=Span(start, end), reason=reason)


def _invalid_declaration_identifier_start(source: str, span: Span) -> _InvalidIdentifierStartHit | None:
    return _invalid_identifier_start_at(source, span, statement_tokens(source, span), 0)


def _invalid_parameter_identifier_start(source: str, span: Span) -> _InvalidIdentifierStartHit | None:
    toks = statement_tokens(source, span)
    i = 0
    while _is_parameter_modifier(_at(toks, i)):
        i += 1
    return _invalid_identifier_start_at(source, span, toks, i)


def _invalid_procedure_identifier_start(source: str, proc: ProcedureNode) -> _InvalidIdentifierStartHit | None:
    header = _first_line_span(source, proc.span)
    toks = statement_tokens(source, header)
    i = 0
    while i < len(toks) and token_text(toks[i]) in _PROC_MODIFIERS:
        i += 1
    head = token_text(_at(toks, i))
    if head == "property":
        i += 2
    elif head == "sub" or head == "function":
        i += 1
    return _invalid_identifier_start_at(source, header, toks, i)


def _invalid_type_or_enum_identifier_start(
    source: str, span: Span, keyword: str
) -> _InvalidIdentifierStartHit | None:
    header = _first_line_span(source, span)
    toks = statement_tokens(source, header)
    i = 0
    if token_text(_at(toks, i)) in ("public", "private"):
        i += 1
    if token_text(_at(toks, i)) == keyword:
        i += 1
    return _invalid_identifier_start_at(source, header, toks, i)


def _invalid_declare_identifier_start(source: str, span: Span) -> _InvalidIdentifierStartHit | None:
    toks = statement_tokens(source, span)
    kind_index = next((i for i, t in enumerate(toks) if token_text(t) in ("sub", "function")), -1)
    return _invalid_identifier_start_at(source, span, toks, kind_index + 1)


def _invalid_const_directive_identifier_start(source: str, span: Span) -> _InvalidIdentifierStartHit | None:
    toks = statement_tokens(source, span)
    if token_text(_at(toks, 1)) == "const":
        return _invalid_identifier_start_at(source, span, toks, 2)
    return None


def check_invalid_identifier_starts(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def report(kind: str, hit: _InvalidIdentifierStartHit | None) -> None:
        if hit is None:
            return
        if hit.reason == "digit":
            push(
                "invalidIdentifierStart",
                f"Invalid {kind} name '{hit.name}': identifiers cannot start with a digit.",
                hit.span,
            )
        elif hit.reason == "underscore":
            push(
                "invalidIdentifierStart",
                f"Invalid {kind} name '{hit.name}': identifiers cannot start with an underscore.",
                hit.span,
            )
        else:
            char = "-" if hit.reason == "hyphen" else "."
            push(
                "invalidIdentifierCharacter",
                f"Invalid {kind} name '{hit.name}': '{char}' is not allowed in an identifier.",
                hit.span,
            )

    def inspect_variable_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            report("variable", _invalid_declaration_identifier_start(source, decl.span))

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_variable_group(member)
        elif isinstance(member, TypeNode):
            report("user-defined type", _invalid_type_or_enum_identifier_start(source, member.span, "type"))
            for field_node in member.fields:
                report("type field", _invalid_declaration_identifier_start(source, field_node.span))
        elif isinstance(member, EnumNode):
            report("enum", _invalid_type_or_enum_identifier_start(source, member.span, "enum"))
            for enum_member in member.members:
                report("enum member", _invalid_declaration_identifier_start(source, enum_member.span))
        elif isinstance(member, DeclareNode):
            report("Declare procedure", _invalid_declare_identifier_start(source, member.span))
        elif isinstance(member, ConditionalDirectiveNode):
            report("conditional compiler constant", _invalid_const_directive_identifier_start(source, member.span))
        elif isinstance(member, ProcedureNode):
            report("procedure", _invalid_procedure_identifier_start(source, member))
            for param in member.params:
                report("parameter", _invalid_parameter_identifier_start(source, param.span))
            for_each_variable_group(member.body, inspect_variable_group, activity)


# -- checkReservedDeclarationNames -----------------------------------------


def _type_or_enum_name_hit(source: str, span: Span, keyword: str) -> NameTokenHit | None:
    header = _first_line_span(source, span)
    toks = statement_tokens(source, header)
    i = 0
    if token_text(_at(toks, i)) in ("public", "private"):
        i += 1
    if token_text(_at(toks, i)) == keyword:
        i += 1
    tok = _at(toks, i)
    name = token_name(tok) if tok is not None else None
    return name_token_hit(header, tok, name) if tok is not None and name else None


def _declare_name_hit(source: str, span: Span) -> NameTokenHit | None:
    toks = statement_tokens(source, span)
    kind_index = next((i for i, t in enumerate(toks) if token_text(t) in ("sub", "function")), -1)
    tok = _at(toks, kind_index + 1) if kind_index >= 0 else None
    name = token_name(tok) if tok is not None else None
    return name_token_hit(span, tok, name) if tok is not None and name else None


def _procedure_name_hit(source: str, proc: ProcedureNode) -> NameTokenHit | None:
    header = _first_line_span(source, proc.span)
    toks = statement_tokens(source, header)
    i = 0
    while i < len(toks) and token_text(toks[i]) in _PROC_MODIFIERS:
        i += 1
    head = token_text(_at(toks, i))
    if head == "property":
        i += 2
    elif head == "sub" or head == "function":
        i += 1
    tok = _at(toks, i)
    name = token_name(tok) if tok is not None else None
    return name_token_hit(header, tok, name) if tok is not None and name else None


def check_reserved_declaration_names(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def report(kind: str, hit: NameTokenHit | None) -> None:
        if hit is None or hit.bracketed or not is_reserved_identifier(hit.name):
            return
        if kind == "type field" and hit.name.lower() == "type":
            return
        push(
            "invalidDeclarationName",
            f"Reserved VBA keyword '{hit.name}' cannot be used as a {kind} name.",
            hit.span,
        )

    def inspect_variable_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            report("variable", declaration_name_hit(source, decl.span, decl.name))

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_variable_group(member)
        elif isinstance(member, TypeNode):
            report("user-defined type", _type_or_enum_name_hit(source, member.span, "type"))
            for field_node in member.fields:
                report("type field", declaration_name_hit(source, field_node.span, field_node.name))
        elif isinstance(member, EnumNode):
            report("enum", _type_or_enum_name_hit(source, member.span, "enum"))
            for enum_member in member.members:
                report("enum member", declaration_name_hit(source, enum_member.span, enum_member.name))
        elif isinstance(member, DeclareNode):
            report("Declare procedure", _declare_name_hit(source, member.span))
        elif isinstance(member, ProcedureNode):
            report("procedure", _procedure_name_hit(source, member))
            for param in member.params:
                report("parameter", declaration_name_hit(source, param.span, param.name))
            for_each_variable_group(member.body, inspect_variable_group, activity)


# -- checkDimInitializer ---------------------------------------------------


def _top_level_assign_offset(source: str, span: Span) -> int | None:
    toks = [
        t
        for t in tokenize(source[span.start : span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    depth = 0
    for t in toks:
        r = t.raw_text
        if r == "(":
            depth += 1
        elif r == ")":
            depth -= 1
        elif depth == 0 and t.kind is TokenKind.OPERATOR and r == "=":
            return span.start + t.start
    return None


def check_dim_initializer(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def inspect(group: VariableGroupNode) -> None:
        if group.is_const:
            return  # Const requires '='; not an error.
        at = _top_level_assign_offset(source, group.span)
        if at is not None:
            push(
                "dimInitializer",
                "A variable declaration cannot include an initializer in VBA; "
                "assign the value in a separate statement.",
                Span(at, at + 1),
            )

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect(member)
        elif isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, inspect, activity)


# -- checkUnexpectedDeclarationTokens --------------------------------------


def _is_declaration_type_name_token(tok: VbaToken | None) -> bool:
    return tok is not None and tok.kind in (
        TokenKind.IDENTIFIER,
        TokenKind.KEYWORD,
        TokenKind.BRACKETED_IDENTIFIER,
    )


def _consume_declaration_type_name(toks: Sequence[VbaToken], start: int) -> int:
    if not _is_declaration_type_name_token(_at(toks, start)):
        return start
    i = start + 1
    while True:
        dot = _at(toks, i)
        if dot is None or dot.raw_text != ".":
            return i
        if not _is_declaration_type_name_token(_at(toks, i + 1)):
            return start
        i += 2


def _unexpected_token_after_declaration_type(
    source: str, span: Span, allow_equals: bool
) -> tuple[str, Span] | None:
    toks = statement_tokens(source, span)
    as_index = next((i for i, t in enumerate(toks) if token_text(t) == "as"), -1)
    if as_index < 0:
        return None
    i = as_index + 1
    if token_text(_at(toks, i)) == "new":
        i += 1
    type_start = i
    i = _consume_declaration_type_name(toks, i)
    if i == type_start:
        return None
    fixed = parse_fixed_length_string_type(toks, type_start)
    if fixed is not None and fixed.end_index > i:
        i = fixed.end_index
    nxt = _at(toks, i)
    if nxt is None:
        return None
    if allow_equals and nxt.kind is TokenKind.OPERATOR and nxt.raw_text == "=":
        return None
    return (nxt.raw_text, absolute_span(span, nxt))


def _parameter_array_as_type_syntax_hit(source: str, param: ParameterNode) -> tuple[Span, str] | None:
    toks = statement_tokens(source, param.span)
    as_index = next((i for i, t in enumerate(toks) if token_text(t) == "as"), -1)
    if as_index < 0:
        return None
    type_start = as_index + 1
    if token_text(_at(toks, type_start)) == "new":
        type_start += 1
    type_end = _consume_declaration_type_name(toks, type_start)
    if type_end == type_start:
        return None
    open_tok = _at(toks, type_end)
    close_tok = _at(toks, type_end + 1)
    if open_tok is None or close_tok is None or open_tok.raw_text != "(" or close_tok.raw_text != ")":
        return None
    type_name = source[param.span.start + toks[type_start].start : param.span.start + toks[type_end - 1].end]
    return (Span(param.span.start + open_tok.start, param.span.start + close_tok.end), type_name)


def check_unexpected_declaration_tokens(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def inspect(span: Span, allow_equals: bool) -> None:
        hit = _unexpected_token_after_declaration_type(source, span, allow_equals)
        if hit is None:
            return
        text, hit_span = hit
        push(
            "unexpectedDeclarationToken",
            f"Unexpected token '{text}' after a complete declaration type; "
            "this will fail to compile as a syntax error.",
            hit_span,
        )

    def inspect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            inspect(decl.span, True)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, TypeNode):
            for field_node in member.fields:
                inspect(field_node.span, False)
        elif isinstance(member, ProcedureNode):
            for param in member.params:
                if _parameter_array_as_type_syntax_hit(source, param) is None:
                    inspect(param.span, True)
            for_each_variable_group(member.body, inspect_group, activity)


# -- checkTypeDeclarationCharacterAsClause ---------------------------------


def check_type_declaration_character_as_clause(mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def report(node: _TypeDeclarationSuffixNode, label: str) -> None:
        if not node.type_suffix or not node.has_as_clause:
            return
        push(
            "typeDeclarationCharacterAsClause",
            f"{label} '{node.name}' combines type-declaration character '{node.type_suffix}' "
            "with an As clause; use only one type declaration form.",
            node.type_suffix_span if node.type_suffix_span is not None else node.span,
        )

    def inspect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            report(decl, "Const declaration" if group.is_const else "Declaration")

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, TypeNode):
            for field_node in member.fields:
                report(field_node, "Type field")
        elif isinstance(member, ProcedureNode):
            if member.proc_kind is ProcKind.FUNCTION:
                report(member, "Function")
            for param in member.params:
                report(param, "Parameter")
            for_each_variable_group(member.body, inspect_group, activity)


# -- checkOptionPlacement --------------------------------------------------


def check_option_placement(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    declaration_seen = False
    for member in active_module_members(mod, activity):
        if isinstance(member, AttributeNode):
            continue
        if isinstance(member, OptionNode):
            if declaration_seen:
                push(
                    "optionAfterDeclaration",
                    "Option statements must appear before any declaration or procedure.",
                    first_token_span(source, member.span),
                )
            continue
        declaration_seen = True


# -- checkEmptyType --------------------------------------------------------


def check_empty_type(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, TypeNode) or not member.closed:
            continue
        if any(not is_inactive_node(activity, field_node) for field_node in member.fields):
            continue
        push(
            "emptyType",
            f"Type '{member.name}' must declare at least one member.",
            member.name_span if member.name_span is not None else member.span,
        )


# -- checkDuplicateOptions -------------------------------------------------


def check_duplicate_options(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    seen: set[str] = set()
    for member in active_module_members(mod, activity):
        if not isinstance(member, OptionNode):
            continue
        if activity is not None and activity.activity_for_span(member.span) is not ConditionalActivity.ACTIVE:
            continue
        parts = member.option_text.strip().split()
        first_word = parts[0] if parts else ""
        category = first_word.lower()
        if not category:
            continue
        if category in seen:
            push(
                "duplicateOption",
                f"Duplicate Option statement; only one 'Option {first_word}' is allowed per module.",
                first_token_span(source, member.span),
            )
        else:
            seen.add(category)


# -- checkTooManyParameters ------------------------------------------------


def check_too_many_parameters(mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or len(member.params) <= _MAX_PROCEDURE_PARAMETERS:
            continue
        push(
            "tooManyParameters",
            f"A procedure may have at most {_MAX_PROCEDURE_PARAMETERS} parameters; "
            f"'{member.name}' declares {len(member.params)}.",
            member.name_span if member.name_span is not None else member.span,
        )


# -- checkIdentifierTooLong ------------------------------------------------


def check_identifier_too_long(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def report(name: str, span: Span) -> None:
        if len(name) <= _MAX_IDENTIFIER_LENGTH:
            return
        push(
            "identifierTooLong",
            f"Identifier '{name[:24]}...' is {len(name)} characters; "
            f"VBA allows at most {_MAX_IDENTIFIER_LENGTH}.",
            span,
        )

    def inspect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            report(decl.name, decl.name_span if decl.name_span is not None else decl.span)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, TypeNode):
            report(member.name, member.name_span if member.name_span is not None else member.span)
            for field_node in member.fields:
                report(field_node.name, field_node.name_span if field_node.name_span is not None else field_node.span)
        elif isinstance(member, EnumNode):
            report(member.name, member.name_span if member.name_span is not None else member.span)
            for enum_member in member.members:
                report(
                    enum_member.name,
                    enum_member.name_span if enum_member.name_span is not None else enum_member.span,
                )
        elif isinstance(member, ProcedureNode):
            report(member.name, member.name_span if member.name_span is not None else member.span)
            for param in member.params:
                report(param.name, param.name_span if param.name_span is not None else param.span)
            for_each_variable_group(member.body, inspect_group, activity)


# -- checkUdtParameterConstraints ------------------------------------------


def check_udt_parameter_constraints(mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    udt_names: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, TypeNode):
            udt_names.add(member.name.strip().lower())
    if not udt_names:
        return
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        for param in member.params:
            if not param.as_type or param.as_type.strip().lower() not in udt_names:
                continue
            if param.optional:
                push(
                    "optionalUdtParameter",
                    f"Optional parameter '{param.name}' cannot be a user-defined type ('{param.as_type}').",
                    param.name_span if param.name_span is not None else param.span,
                )
            elif param.by_val:
                push(
                    "byvalUdtParameter",
                    f"User-defined type parameter '{param.name}' ('{param.as_type}') "
                    "cannot be passed ByVal; pass it ByRef.",
                    param.name_span if param.name_span is not None else param.span,
                )
