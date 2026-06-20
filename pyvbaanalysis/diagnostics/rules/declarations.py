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

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal, Union

from ...completion.member_access import MemberCompletionContext
from ...completion.type_completion import (
    TypeCompletionKind,
    is_creatable_type_completion,
    resolve_type_name,
)
from ...conditional import (
    ConditionalActivity,
    ConditionalActivityTracker,
    collect_conditional_directives,
)
from ...lexer.keyword_table import OPERATOR_IDENTIFIERS, is_reserved_identifier
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.fixed_length_string import parse_fixed_length_string_type
from ...parser.nodes import (
    AttributeNode,
    BodyNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    DeclareNode,
    EnumNode,
    EventNode,
    LeafStatementNode,
    ModuleMember,
    ModuleNode,
    OptionNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    TypeFieldNode,
    TypeNode,
    VariableDeclNode,
    VariableGroupNode,
)
from ...parser.type_declaration_suffix import is_type_declaration_suffix
from ...runtime import resolve_runtime_function
from ...types.type_names import (
    is_known_object_assignment_type,
    is_known_scalar_type,
    normalize_type,
)
from ..const_expr import (
    collect_body_literal_integer_constants,
    collect_module_literal_integer_constants,
    resolve_fixed_length_string_size,
)
from ..context import AnalyzeModuleOptions, PushFn, statement_tokens
from ..walker import (
    absolute_span,
    active_module_members,
    declared_name_span,
    first_token_span,
    for_each_body_statement,
    for_each_variable_group,
    is_inactive_node,
    match_paren_from,
    pluralize_count,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
    token_text,
    top_level_operator_index,
)
from .shared import (
    DEFTYPE_KEYWORDS,
    NameTokenHit,
    declaration_name_hit,
    leading_declaration_modifier_count,
    module_declaration_statement_in_procedure,
    name_token_hit,
    scan_conditional_compilation_branch_order,
)
# resolveKnownObjectAssignmentType lives in XLIDE's shared typeInference.ts; the
# port keeps it next to its other consumer (typeOfIs) and shares it from there.
from .type_of_is import _resolve_known_object_assignment_type

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


# -- checkPropertySetterValueParameters ------------------------------------

# Object-value branch (propertyLetObjectValue) is DEFERRED: it needs
# resolveKnownObjectAssignmentType (host/project class resolution), which the port
# does not yet expose as a richer typed result. The three branches below are pure
# signature/structure checks and are sound without that surface.


def _property_setter_return_type_span(source: str, proc: ProcedureNode) -> Span:
    """Span of the offending `As <type>` return clause on a Property Let/Set header.

    Port of propertySetterReturnTypeSpan: walk past modifiers, `Property`, the
    accessor, the name, and any parameter list, to the `As` keyword and its type.
    Falls back to the `As` keyword span when the layout is unexpected.
    """
    header = _first_line_span(source, proc.span)
    toks = statement_tokens(source, header)
    i = 0
    while i < len(toks) and token_text(toks[i]) in _PROC_MODIFIERS:
        i += 1
    if token_text(_at(toks, i)) == "property":
        i += 2  # Property + Let/Set
    i += 1  # property name
    open_tok = _at(toks, i)
    if open_tok is None or open_tok.raw_text != "(":
        return _keyword_span(source, header, "as")
    depth = 0
    while i < len(toks):
        raw = toks[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
            if depth == 0:
                i += 1
                break
        i += 1
    as_tok = _at(toks, i)
    if as_tok is None or token_text(as_tok) != "as":
        return _keyword_span(source, header, "as")
    type_start = i + 1
    type_end = _consume_declaration_type_name(toks, type_start)
    if type_end == type_start:
        type_end = i + 1
    end_tok = _at(toks, type_end - 1) or as_tok
    return Span(header.start + as_tok.start, header.start + end_tok.end)


def check_property_setter_value_parameters(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, member_ctx: MemberCompletionContext, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind not in (
            ProcKind.PROPERTY_LET,
            ProcKind.PROPERTY_SET,
        ):
            continue
        label = "Property Let" if member.proc_kind is ProcKind.PROPERTY_LET else "Property Set"
        if member.has_as_clause:
            push(
                "propertySetterReturnType",
                f"{label} '{member.name}' cannot declare a return type; "
                "use the final value parameter for the assigned value.",
                _property_setter_return_type_span(source, member),
            )
        if len(member.params) > 0:
            value_param = member.params[-1]
            if member.proc_kind is ProcKind.PROPERTY_SET:
                normalized = normalize_type(value_param.as_type)
                if normalized and is_known_scalar_type(normalized):
                    push(
                        "propertySetScalarValue",
                        f"Property Set '{member.name}' final value parameter "
                        f"'{value_param.name}' must be an object reference, but it is "
                        f"declared As {value_param.as_type}.",
                        declared_name_span(source, value_param.span, value_param.name),
                    )
            else:
                object_type = _resolve_known_object_assignment_type(
                    value_param.as_type, member_ctx
                )
                if object_type is not None:
                    push(
                        "propertyLetObjectValue",
                        f"Property Let '{member.name}' final value parameter "
                        f"'{value_param.name}' must not be an object reference; use "
                        f"Property Set because it is declared As {object_type.display}.",
                        declared_name_span(source, value_param.span, value_param.name),
                    )
            continue
        push(
            "propertySetterMissingValue",
            f"{label} '{member.name}' must include a final value parameter.",
            declared_name_span(source, member.span, member.name),
        )


# -- checkInvalidAsTypeNames (safe branches) -------------------------------
#
# Full port of checkInvalidAsTypeNames / collectTypeNameReferences. Each type-name
# reference (As-clause, As New, New expression, TypeOf...Is, Implements, return
# type) is resolved with resolveTypeName over the project-type registry
# (opts.project_types) + host model. A name that resolves to the 'ambiguous' marker
# (multiple visible project types share it), a New reference to a non-creatable
# non-host type, a reserved VBA identifier, a VBA runtime function, or a known
# project non-type declaration is reported; anything that resolves to a real type is
# accepted. Qualified references (`Mod.Type`) resolve through the module-qualified
# candidate set. The no-false-positive guarantee comes from resolveTypeName itself
# (faithful to XLIDE) plus the project-type context the caller threads in.

_TypeNameReferenceKind = Literal[
    "declaration", "newDeclaration", "newExpression", "typeOfIs", "implements"
]


@dataclass(frozen=True, slots=True)
class _TypeNameRef:
    name: str
    span: Span
    kind: _TypeNameReferenceKind
    qualifier: str | None = None


def _type_reference_lookup_name(ref: _TypeNameRef) -> str:
    """Port of typeReferenceLookupName: qualified refs look up as `Qualifier.Member`."""
    return f"{ref.qualifier}.{ref.name}" if ref.qualifier else ref.name


def _type_name_ref_from_tokens(
    toks: Sequence[VbaToken], type_index: int, base: int, kind: _TypeNameReferenceKind
) -> _TypeNameRef | None:
    first = _at(toks, type_index)
    if first is None:
        return None
    first_name = token_name(first)
    if not first_name:
        return None
    dot = _at(toks, type_index + 1)
    member_tok = _at(toks, type_index + 2) if dot is not None and dot.raw_text == "." else None
    if member_tok is None:
        return _TypeNameRef(name=first_name, span=Span(base + first.start, base + first.end), kind=kind)
    member_name = token_name(member_tok)
    if not member_name:
        return _TypeNameRef(name=first_name, span=Span(base + first.start, base + first.end), kind=kind)
    return _TypeNameRef(
        name=member_name,
        span=Span(base + member_tok.start, base + member_tok.end),
        kind=kind,
        qualifier=first_name,
    )


def _type_name_after_as(source: str, span: Span) -> _TypeNameRef | None:
    toks = statement_tokens(source, span)
    for i, tok in enumerate(toks):
        if token_text(tok) != "as":
            continue
        type_index = i + 1
        kind: _TypeNameReferenceKind = "declaration"
        if token_text(_at(toks, type_index)) == "new":
            type_index += 1
            kind = "newDeclaration"
        ref = _type_name_ref_from_tokens(toks, type_index, span.start, kind)
        if ref is not None:
            return ref
    return None


def _return_type_name_ref(source: str, proc: ProcedureNode) -> _TypeNameRef | None:
    header = _first_line_span(source, proc.span)
    toks = statement_tokens(source, header)
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
            continue
        if raw == ")":
            depth -= 1
            continue
        if depth != 0 or token_text(tok) != "as":
            continue
        return _type_name_ref_from_tokens(toks, i + 1, header.start, "declaration")
    return None


def _type_names_after_new(source: str, span: Span) -> list[_TypeNameRef]:
    toks = statement_tokens(source, span)
    out: list[_TypeNameRef] = []
    for i, tok in enumerate(toks):
        if token_text(tok) != "new":
            continue
        ref = _type_name_ref_from_tokens(toks, i + 1, span.start, "newExpression")
        if ref is not None:
            out.append(ref)
    return out


def _type_names_after_typeof_is(source: str, span: Span) -> list[_TypeNameRef]:
    toks = statement_tokens(source, span)
    out: list[_TypeNameRef] = []
    saw_typeof = False
    for i, tok in enumerate(toks):
        lower = token_text(tok)
        if lower == "typeof":
            saw_typeof = True
            continue
        if not saw_typeof or lower != "is":
            continue
        ref = _type_name_ref_from_tokens(toks, i + 1, span.start, "typeOfIs")
        if ref is not None:
            out.append(ref)
        saw_typeof = False
    return out


_IMPLEMENTS_RE = re.compile(
    r"^\s*Implements\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\b",
    re.IGNORECASE,
)


def _collect_implements_refs(source: str) -> list[_TypeNameRef]:
    out: list[_TypeNameRef] = []
    line_start = 0
    length = len(source)
    while line_start <= length:
        line_end = source.find("\n", line_start)
        if line_end < 0:
            line_end = length
        line = source[line_start:line_end]
        if line.endswith("\r"):
            line = line[:-1]
        code = re.sub(r"'.*$", "", line)
        match = _IMPLEMENTS_RE.match(code)
        if match:
            raw_name = match.group(1)
            column = line.find(raw_name, match.start())
            dot = raw_name.find(".")
            if column >= 0 and dot > 0:
                out.append(
                    _TypeNameRef(
                        name=raw_name[dot + 1 :],
                        span=Span(line_start + column + dot + 1, line_start + column + len(raw_name)),
                        kind="implements",
                        qualifier=raw_name[:dot],
                    )
                )
            elif column >= 0:
                out.append(
                    _TypeNameRef(
                        name=raw_name,
                        span=Span(line_start + column, line_start + column + len(raw_name)),
                        kind="implements",
                    )
                )
        if line_end == length:
            break
        line_start = line_end + 1
    return out


def _collect_type_name_references(source: str, mod: ModuleNode) -> list[_TypeNameRef]:
    out: list[_TypeNameRef] = []
    out.extend(_collect_implements_refs(source))

    def collect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            if decl.as_type:
                ref = _type_name_after_as(source, decl.span)
                if ref is not None:
                    out.append(ref)

    def collect_statement(span: Span) -> None:
        out.extend(_type_names_after_new(source, span))
        out.extend(_type_names_after_typeof_is(source, span))

    def collect_body(body: Sequence[BodyNode]) -> None:
        for_each_variable_group(list(body), collect_group)
        for_each_body_statement(list(body), lambda stmt: collect_statement(stmt.span))

    for member in mod.members:
        if isinstance(member, VariableGroupNode):
            collect_group(member)
        elif isinstance(member, TypeNode):
            for field_node in member.fields:
                if field_node.as_type:
                    ref = _type_name_after_as(source, field_node.span)
                    if ref is not None:
                        out.append(ref)
        elif isinstance(member, ProcedureNode):
            for param in member.params:
                if param.as_type:
                    ref = _type_name_after_as(source, param.span)
                    if ref is not None:
                        out.append(ref)
            if member.return_type:
                ref = _return_type_name_ref(source, member)
                if ref is not None:
                    out.append(ref)
            collect_body(member.body)

    out.sort(key=lambda ref: (ref.span.start, ref.span.end))
    return out


_TYPE_KIND_LABEL_FOR_NEW: dict[TypeCompletionKind, str] = {
    "primitive": "a VBA primitive type",
    "external": "an external interface type",
    "host": "an Excel object-model type",
    "document": "a document module type",
    "enum": "an Enum type",
    "userType": "a user-defined Type",
    "ambiguous": "an ambiguous project type",
    "module": "a module qualifier",
    "class": "a creatable project type",
    "userform": "a creatable project type",
}


def _is_new_type_reference(kind: _TypeNameReferenceKind) -> bool:
    return kind == "newExpression" or kind == "newDeclaration"


def _collect_with_events_new_declaration_spans(
    mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> list[Span]:
    """Port of collectWithEventsNewDeclarationSpans: the spans of `WithEvents x As
    New T` declarations. `New` on a WithEvents declaration is legal (the field is
    initialized lazily), so its New reference is exempt from invalidNewTypeName."""
    spans: list[Span] = []

    def inspect(group: VariableGroupNode) -> None:
        if not group.with_events or is_inactive_node(activity, group):
            return
        for decl in group.declarations:
            if decl.is_new:
                spans.append(decl.span)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect(member)
            continue
        if isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, inspect, activity)
    return spans


def check_invalid_as_type_names(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, opts: AnalyzeModuleOptions, push: PushFn) -> None:
    with_events_new_spans = _collect_with_events_new_declaration_spans(mod, activity)
    known_non_type_names = opts.known_non_type_names or frozenset()
    for ref in _collect_type_name_references(source, mod):
        if activity is not None and activity.is_inactive(ref.span):
            continue
        lookup_name = _type_reference_lookup_name(ref)
        resolved = resolve_type_name(lookup_name, opts.project_types, opts.host_model)
        if resolved is not None and resolved.kind == "ambiguous":
            push(
                "invalidAsTypeName",
                f"'{ref.name}' is ambiguous because multiple visible project types use that name.",
                ref.span,
            )
            continue
        if (
            resolved is not None
            and _is_new_type_reference(ref.kind)
            and not is_creatable_type_completion(resolved)
            and resolved.kind != "host"
        ):
            if ref.kind == "newDeclaration" and any(
                _contains_span(span, ref.span) for span in with_events_new_spans
            ):
                continue
            push(
                "invalidNewTypeName",
                f"'{ref.name}' is {_TYPE_KIND_LABEL_FOR_NEW[resolved.kind]} and cannot be used "
                "with New. New can create project classes and UserForms only.",
                ref.span,
            )
            continue
        if resolved is not None:
            continue
        if is_reserved_identifier(ref.name):
            push(
                "invalidAsTypeName",
                f"'{ref.name}' is a reserved VBA identifier, not a valid type name.",
                ref.span,
            )
            continue
        if resolve_runtime_function(ref.name) is not None:
            push(
                "invalidAsTypeName",
                f"'{ref.name}' is a VBA runtime function, not a valid type name.",
                ref.span,
            )
            continue
        if ref.name.lower() in known_non_type_names:
            push(
                "invalidAsTypeName",
                f"'{ref.name}' resolves to a project declaration, but that declaration is not a type.",
                ref.span,
            )
            continue


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


# -- checkFixedLengthStringBounds ------------------------------------------

# MS-VBAL fixed-length String bounds (VBE oracle: "Invalid length for fixed-length
# string"). The active rule resolves decimal integer literal sizes and same
# module/procedure Const aliases that reduce to a decimal integer literal; unknown,
# duplicate, string, and compound constants stay deferred until broader
# constant-expression semantics are verified.
_FIXED_LENGTH_STRING_MIN = 1
_FIXED_LENGTH_STRING_MAX = 65526


def check_fixed_length_string_bounds(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    module_constants = collect_module_literal_integer_constants(mod, activity)

    def inspect_declaration(decl: VariableDeclNode | TypeFieldNode, constants: dict[str, int | None]) -> None:
        if decl.fixed_length is None or is_inactive_node(activity, decl):
            return
        value = resolve_fixed_length_string_size(decl.fixed_length, constants)
        if value is None or _FIXED_LENGTH_STRING_MIN <= value <= _FIXED_LENGTH_STRING_MAX:
            return
        push(
            "fixedLengthStringSize",
            f"Fixed-length String size must be between {_FIXED_LENGTH_STRING_MIN} and "
            f"{_FIXED_LENGTH_STRING_MAX} characters; got {value}.",
            _fixed_length_string_length_span(source, decl.span) or decl.span,
        )

    def inspect_group(group: VariableGroupNode) -> None:
        for decl in group.declarations:
            inspect_declaration(decl, module_constants)

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, TypeNode):
            for field_node in member.fields:
                inspect_declaration(field_node, module_constants)
        elif isinstance(member, ProcedureNode):
            procedure_constants = dict(module_constants)
            collect_body_literal_integer_constants(member.body, procedure_constants, activity)
            body_groups: list[VariableGroupNode] = []
            for_each_variable_group(member.body, body_groups.append, activity)
            for group in body_groups:
                for decl in group.declarations:
                    inspect_declaration(decl, procedure_constants)


def _fixed_length_string_length_span(source: str, span: Span) -> Span | None:
    toks = statement_tokens(source, span)
    as_index = next((i for i, tok in enumerate(toks) if token_text(tok) == "as"), -1)
    if as_index < 0:
        return None
    type_start = as_index + 1
    if type_start < len(toks) and token_text(toks[type_start]) == "new":
        type_start += 1
    fixed = parse_fixed_length_string_type(toks, type_start)
    if fixed is None:
        return None
    return absolute_span(span, toks[fixed.length_index])


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


# -- checkParameterOrder ---------------------------------------------------


def check_parameter_order(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        params = member.params
        has_optional = any(p.optional for p in params)
        optional_seen = False
        for i, p in enumerate(params):
            array_as_type = _parameter_array_as_type_syntax_hit(source, p)
            if array_as_type is not None:
                array_span, type_name = array_as_type
                push(
                    "parameterArrayAsTypeSyntax",
                    f"Array parameter '{p.name}' must place parentheses after the parameter name, "
                    f"before the As clause; use '{p.name}() As {type_name}'.",
                    array_span,
                )
                if p.optional:
                    optional_seen = True
                continue
            if p.param_array:
                if p.as_type and normalize_type(p.as_type) != "variant":
                    push(
                        "paramArrayNonVariant",
                        f"ParamArray '{p.name}' elements must be Variant, but this parameter "
                        f"is declared As {p.as_type}.",
                        declared_name_span(source, p.span, p.name),
                    )
                if has_optional:
                    push(
                        "paramArrayWithOptional",
                        f"ParamArray '{p.name}' cannot be used in the same parameter list as "
                        "Optional arguments.",
                        declared_name_span(source, p.span, p.name),
                    )
                if i != len(params) - 1:
                    push(
                        "paramArrayNotLast",
                        f"ParamArray '{p.name}' must be the last parameter.",
                        declared_name_span(source, p.span, p.name),
                    )
                continue
            if p.optional:
                optional_seen = True
                continue
            if optional_seen:
                push(
                    "requiredParamAfterOptional",
                    f"Parameter '{p.name}' must be Optional because it follows an Optional parameter.",
                    declared_name_span(source, p.span, p.name),
                )


# -- checkPropertyAccessorSignatures ---------------------------------------

_PROPERTY_KINDS = (ProcKind.PROPERTY_GET, ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET)


@dataclass(slots=True)
class _PropertyAccessorGroup:
    name: str
    gets: list[ProcedureNode]
    setters: list[ProcedureNode]


def _effective_passing_mode(param: ParameterNode) -> str:
    return "byval" if param.by_val else "byref"


def _property_procedure_label(kind: ProcKind) -> str:
    if kind is ProcKind.PROPERTY_GET:
        return "Property Get"
    if kind is ProcKind.PROPERTY_LET:
        return "Property Let"
    if kind is ProcKind.PROPERTY_SET:
        return "Property Set"
    return "Property"


def _property_parameter_type_mismatch(
    expected: ParameterNode, actual: ParameterNode, index: int
) -> str | None:
    expected_type = normalize_type(expected.as_type) or "variant"
    actual_type = normalize_type(actual.as_type) or "variant"
    if expected_type == actual_type:
        return None
    scalar_or_variant = (expected_type == "variant" or is_known_scalar_type(expected_type)) and (
        actual_type == "variant" or is_known_scalar_type(actual_type)
    )
    if not scalar_or_variant:
        return None
    return (
        f"Index parameter {index} type must match: expected {expected.as_type or 'Variant'}, "
        f"found {actual.as_type or 'Variant'}."
    )


def _property_index_parameter_mismatch(
    get_params: list[ParameterNode], setter_index_params: list[ParameterNode]
) -> str | None:
    if len(get_params) != len(setter_index_params):
        return (
            f"Expected {pluralize_count(len(get_params), 'index parameter')}, "
            f"but found {len(setter_index_params)}."
        )
    for i in range(len(get_params)):
        expected = get_params[i]
        actual = setter_index_params[i]
        if expected.is_array != actual.is_array:
            return f"Index parameter {i + 1} array shape must match."
        if _effective_passing_mode(expected) != _effective_passing_mode(actual):
            return f"Index parameter {i + 1} passing mode must match."
        type_reason = _property_parameter_type_mismatch(expected, actual, i + 1)
        if type_reason:
            return type_reason
    return None


def check_property_accessor_signatures(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    groups: dict[str, _PropertyAccessorGroup] = {}
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode) or member.proc_kind not in _PROPERTY_KINDS:
            continue
        key = member.name.lower()
        group = groups.get(key)
        if group is None:
            group = _PropertyAccessorGroup(name=member.name, gets=[], setters=[])
            groups[key] = group
        if member.proc_kind is ProcKind.PROPERTY_GET:
            group.gets.append(member)
        else:
            group.setters.append(member)

    for group in groups.values():
        if len(group.gets) != 1:
            continue
        getter = group.gets[0]
        for setter in group.setters:
            if len(setter.params) == 0:
                continue
            reason = _property_index_parameter_mismatch(getter.params, setter.params[:-1])
            if reason is None:
                continue
            push(
                "propertyAccessorSignatureMismatch",
                f"{_property_procedure_label(setter.proc_kind)} '{setter.name}' argument list "
                f"must match Property Get '{getter.name}' before the final value parameter. {reason}",
                declared_name_span(source, setter.span, setter.name),
            )


# -- checkNonConstantConstValues / checkNonConstantEnumMemberValues --------

# Operator keywords (And, Or, Not, Mod, ...) lex as keyword but are never callable
# names; exclude them from the call heuristic so `6 And (3)` is not read as a call.
_OPERATOR_KEYWORD_WORDS: frozenset[str] = frozenset(w.lower() for w in OPERATOR_IDENTIFIERS)


def _non_constant_default_element(
    tokens: list[VbaToken], base_offset: int
) -> tuple[str, Span] | None:
    for i, tok in enumerate(tokens):
        word = (tok.canonical_text if tok.canonical_text is not None else tok.raw_text).lower()
        if tok.kind is TokenKind.KEYWORD and (word == "new" or word == "addressof"):
            return (f"'{tok.raw_text}'", Span(base_offset + tok.start, base_offset + tokens[-1].end))
        is_name = tok.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED_IDENTIFIER)
        is_operator_keyword = tok.kind is TokenKind.KEYWORD and word in _OPERATOR_KEYWORD_WORDS
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        if is_name and not is_operator_keyword and nxt is not None and nxt.raw_text == "(":
            close_index = match_paren_from(tokens, i + 1)
            end_tok = tokens[close_index] if close_index >= 0 else tokens[i + 1]
            return (f"the call '{tok.raw_text}(...)'", Span(base_offset + tok.start, base_offset + end_tok.end))
    return None


def _value_tokens_after_equals(source: str, span: Span) -> list[VbaToken] | None:
    toks = [
        t
        for t in tokenize(source[span.start : span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    eq = top_level_operator_index(toks, "=")
    if eq < 0 or eq + 1 >= len(toks):
        return None
    return toks[eq + 1 :]


def check_non_constant_parameter_defaults(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """An Optional parameter default must be a constant expression (no call/New/AddressOf).

    Object-typed parameters are skipped: their defaults are owned by the
    parameter-default-type-mismatch rule (\"must be Nothing\"). Host-class object
    typing resolves to None here (M9), so those parameters are still scanned —
    sound, since a non-constant default is invalid regardless of the object type.
    """
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        for param in member.params:
            if not param.default_raw or is_known_object_assignment_type(param.as_type):
                continue
            tokens = _value_tokens_after_equals(source, param.span)
            if tokens is None:
                continue
            non_constant = _non_constant_default_element(tokens, param.span.start)
            if non_constant is None:
                continue
            label, hit_span = non_constant
            push(
                "parameterDefaultNotConstant",
                f"Optional parameter '{param.name}' default must be a constant expression; "
                f"{label} is not constant.",
                hit_span,
            )


def check_non_constant_const_values(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def inspect_group(group: VariableGroupNode) -> None:
        if not group.is_const:
            return
        for decl in group.declarations:
            if decl.default_raw is None or is_inactive_node(activity, decl):
                continue
            tokens = _value_tokens_after_equals(source, decl.span)
            if tokens is None:
                continue
            non_constant = _non_constant_default_element(tokens, decl.span.start)
            if non_constant is None:
                continue
            label, hit_span = non_constant
            push(
                "constValueNotConstant",
                f"Const '{decl.name}' value must be a constant expression; {label} is not constant.",
                hit_span,
            )

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            inspect_group(member)
        elif isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, inspect_group, activity)


def check_non_constant_enum_member_values(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, EnumNode):
            continue
        for enum_member in member.members:
            if enum_member.value_raw is None or is_inactive_node(activity, enum_member):
                continue
            tokens = _value_tokens_after_equals(source, enum_member.span)
            if tokens is None:
                continue
            non_constant = _non_constant_default_element(tokens, enum_member.span.start)
            if non_constant is None:
                continue
            label, hit_span = non_constant
            push(
                "enumMemberNotConstant",
                f"Enum member '{enum_member.name}' value must be a constant expression; "
                f"{label} is not constant.",
                hit_span,
            )


# -- module-declaration placement rules ------------------------------------


def _contains_span(container: Span, inner: Span) -> bool:
    return inner.start >= container.start and inner.end <= container.end


def _keyword_span(source: str, span: Span, *keywords: str) -> Span:
    expected = set(keywords)
    for tok in statement_tokens_after_leading_label(source, span):
        if token_text(tok) in expected:
            return absolute_span(span, tok)
    return first_token_span(source, span)


def _deftype_module_declaration_hit(source: str, span: Span) -> tuple[str, Span] | None:
    toks = statement_tokens_after_leading_label(source, span)
    first = toks[0] if toks else None
    if first is None or token_text(first) not in DEFTYPE_KEYWORDS:
        return None
    label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " statements"
    return (label, absolute_span(span, first))


def _module_declaration_after_procedure_hit(source: str, member: ModuleMember) -> tuple[str, Span] | None:
    if isinstance(member, DeclareNode):
        return ("Declare statements", _keyword_span(source, member.span, "declare"))
    if isinstance(member, EventNode):
        return ("Event declarations", _keyword_span(source, member.span, "event"))
    if isinstance(member, VariableGroupNode):
        if member.is_const:
            return ("Const declarations", _keyword_span(source, member.span, "const"))
        return ("Module variable declarations", first_token_span(source, member.span))
    if isinstance(member, TypeNode):
        return ("Type declarations", _keyword_span(source, member.span, "type"))
    if isinstance(member, EnumNode):
        return ("Enum declarations", _keyword_span(source, member.span, "enum"))
    if isinstance(member, StatementNode):
        return _deftype_module_declaration_hit(source, member.span)
    return None


def _is_inside_module_conditional_compilation_block(mod: ModuleNode, span: Span) -> bool:
    depth = 0
    for occ in collect_conditional_directives(mod):
        if occ.container.kind != "module":
            continue
        directive = occ.directive
        if directive.span.start >= span.start:
            break
        if directive.directive_kind is ConditionalDirectiveKind.IF:
            depth += 1
        elif directive.directive_kind is ConditionalDirectiveKind.END_IF:
            depth = max(0, depth - 1)
    return depth > 0


def _module_declaration_after_procedure_message(
    label: str, mod: ModuleNode, member: ModuleMember, activity: ConditionalActivityTracker | None
) -> str:
    if not _is_inside_module_conditional_compilation_block(mod, member.span):
        return f"{label} belong in the module declarations section, before procedures."
    branch_status = activity.activity_for_span(member.span) if activity is not None else None
    if branch_status is ConditionalActivity.ACTIVE:
        return (
            f"{label} in the active conditional-compilation branch belong in the module "
            "declarations section, before procedures."
        )
    return (
        f"{label} in a conditional-compilation branch belong in the module declarations section, "
        "before procedures."
    )


def _is_alternative_procedure_header_statement(source: str, span: Span, procedure: ProcedureNode) -> bool:
    toks = statement_tokens_after_leading_label(source, span)
    i = leading_declaration_modifier_count(toks)
    head = token_text(_at(toks, i))
    kind: ProcKind | None = None
    if head == "property":
        accessor = token_text(_at(toks, i + 1))
        if accessor == "get":
            kind = ProcKind.PROPERTY_GET
        elif accessor == "let":
            kind = ProcKind.PROPERTY_LET
        elif accessor == "set":
            kind = ProcKind.PROPERTY_SET
        i += 2
    elif head == "function":
        kind = ProcKind.FUNCTION
        i += 1
    elif head == "sub":
        kind = ProcKind.SUB
        i += 1
    name_tok = _at(toks, i)
    name = token_name(name_tok) if name_tok is not None else None
    return (
        kind is not None
        and kind == procedure.proc_kind
        and name is not None
        and name.lower() == procedure.name.lower()
    )


def check_module_declarations_in_procedure_bodies(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    def inspect_statement(stmt: LeafStatementNode) -> None:
        hit = module_declaration_statement_in_procedure(source, stmt.span)
        if hit is None:
            return
        label, hit_span = hit
        push(
            "moduleDeclarationInProcedure",
            f"{label} must appear in the module declarations section, not inside a procedure.",
            hit_span,
        )

    def inspect_procedure_body(procedure: ProcedureNode) -> None:
        saw_conditional_directive = False
        for node in procedure.body:
            if isinstance(node, ConditionalDirectiveNode):
                saw_conditional_directive = True
                continue
            if is_inactive_node(activity, node):
                continue
            if isinstance(node, StatementNode):
                if saw_conditional_directive and _is_alternative_procedure_header_statement(source, node.span, procedure):
                    continue
                inspect_statement(node)
                continue
            child = getattr(node, "body", None)
            if isinstance(child, list):
                for_each_body_statement(child, inspect_statement, activity)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            inspect_procedure_body(member)


def check_module_declarations_after_procedures(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    procedure_seen = False
    malformed_conditional_blocks = scan_conditional_compilation_branch_order(mod).malformed_block_spans
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            procedure_seen = True
            continue
        if not procedure_seen:
            continue
        hit = _module_declaration_after_procedure_hit(source, member)
        if hit is None:
            continue
        label, hit_span = hit
        if any(_contains_span(block, member.span) for block in malformed_conditional_blocks):
            continue
        push(
            "moduleDeclarationAfterProcedure",
            _module_declaration_after_procedure_message(label, mod, member, activity),
            hit_span,
        )


def check_module_level_statements_outside_procedures(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    for member in active_module_members(mod, activity):
        if not isinstance(member, StatementNode):
            continue
        toks = statement_tokens_after_leading_label(source, member.span)
        first = toks[0] if toks else None
        if first is None:
            continue
        head = token_text(first)
        if head in DEFTYPE_KEYWORDS or head == "implements":
            continue
        label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " statement"
        push(
            "statementOutsideProcedure",
            f"{label} is invalid outside a Sub, Function, or Property procedure.",
            absolute_span(member.span, first),
        )
