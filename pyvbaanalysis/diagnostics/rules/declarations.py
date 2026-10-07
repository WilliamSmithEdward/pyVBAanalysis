"""Rule family: declaration-site rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/declarations.ts. The
self-contained checks cover procedure headers, identifier spelling, reserved
names, Dim initializers, unexpected declaration tokens, type-declaration
characters, Option placement/duplication, empty Type, parameter/identifier
limits, and UDT parameter constraints. The checks that draw on type inference,
the member-completion context, constant-expression evaluation, or host/runtime
resolution are here too: parameter defaults (non-constant), fixed-length-string
bounds, As-type-name validation, property setter/accessor value types,
non-constant Const/Enum values, and parameter order.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Literal, Union, cast

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
from ...host.library_type_names import library_type_names
from ...js_compat import js_trim
from ...lexer.keyword_table import OPERATOR_IDENTIFIERS, is_reserved_identifier
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.fixed_length_string import parse_fixed_length_string_type
from ...parser.nodes import (
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
from ...symbols.symbol_model import ModuleSymbolKind, VbaProjectTypeKind
from ...types.type_names import is_known_scalar_type, normalize_type
from ..const_expr import (
    collect_body_literal_integer_constants,
    collect_module_literal_integer_constants,
    resolve_fixed_length_string_size,
)
from ..context import AnalyzeModuleOptions, PushFn, is_object_module_kind, statement_tokens
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
    span_for_tokens,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
    token_text,
    top_level_operator_index,
)
from .expressions import juxtaposed_value_index
from .shared import (
    DEFTYPE_KEYWORDS,
    NameTokenHit,
    declaration_name_hit,
    leading_declaration_modifier_count,
    module_declaration_statement_in_procedure,
    name_token_hit,
    report_repeated_keys,
    scan_conditional_compilation_branch_order,
)
# resolveKnownObjectAssignmentType lives in XLIDE's typeInference.ts; the port
# keeps it with the member-access resolver, whose project and host lookups it uses.
from ...completion.member_access import resolve_known_object_assignment_type

# Access/storage modifiers that may lead a procedure declaration.
_PROC_MODIFIERS: frozenset[str] = frozenset({"public", "private", "friend", "global", "static"})
_MAX_PROCEDURE_PARAMETERS = 60
# In a class module the most is 59 (XLIDE issue #210).
_MAX_CLASS_PROCEDURE_PARAMETERS = 59
_MAX_IDENTIFIER_LENGTH = 255

# Nodes that may carry a legacy type-declaration suffix plus an As clause.
_TypeDeclarationSuffixNode = Union[ParameterNode, ProcedureNode, TypeFieldNode, VariableDeclNode]


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


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


# The names the VBE refuses for a module (XLIDE issues #247 and #357, measured in
# Excel 16.0): adding one fails with 0x800AC3D4, renaming to one with
# 50132. Line, Width, Name, Err, Mid, Time, Error, Reset, Beep, Load,
# Unload, Access, Base, Compare, Explicit, Object, Property and Step are
# accepted.
_REFUSED_MODULE_NAMES = frozenset(
    {
        "addressof", "and", "any", "array", "as", "attribute", "boolean", "byref", "byte", "byval",
        "call", "case", "cdate", "circle", "close", "const", "currency", "date", "debug", "decimal",
        "declare", "dim", "do", "double", "each", "else", "elseif", "empty", "end", "enum", "eqv",
        "erase", "event", "exit", "false", "for", "friend", "function", "get", "global", "gosub",
        "goto", "if", "imp", "implements", "in", "input", "integer", "is", "lbound", "len", "lenb",
        "let", "like", "lock", "long", "longlong", "longptr", "loop", "lset", "me", "mod", "new",
        "next", "not", "nothing", "null", "on", "open", "option", "optional", "or", "paramarray",
        "preserve", "print", "private", "pset", "public", "put", "raiseevent", "redim", "rem",
        "resume", "return", "rset", "scale", "seek", "select", "set", "shared", "single", "spc",
        "static", "stop", "string", "sub", "tab", "then", "to", "true", "type", "typeof", "unlock",
        "until", "variant", "wend", "while", "with", "withevents", "write", "xor",
    }
)  # fmt: skip

# The libraries every project of a host references, which a module cannot
# share a name with: renaming one to Excel, VBA, Office or stdole fails
# with 32813, "Name conflicts with existing module, project, or object
# library" (XLIDE issue #357, measured in Excel 16.0). Word, Access and
# PowerPoint are accepted there, being libraries an Excel project does
# not reference.
_REFERENCED_LIBRARIES = frozenset({"vba", "office", "stdole"})
_HOST_LIBRARIES = frozenset({"excel", "word", "powerpoint", "access"})

# `/^[ \t]*Attribute[ \t]+VB_Name[ \t]*=[ \t]*"([^"]*)"/im`: JavaScript's `^`
# under the m flag also follows a lone CR, U+2028 and U+2029, and its i flag
# folds ASCII only.
_VB_NAME_ATTRIBUTE_RE = re.compile(
    '(?:^|(?<=[\\r\\u2028\\u2029]))[ \\t]*Attribute[ \\t]+VB_Name[ \\t]*=[ \\t]*"([^"]*)"',
    re.IGNORECASE | re.MULTILINE | re.ASCII,
)
_FIRST_LINE_END_RE = re.compile(r"\r?\n|\Z")


def check_module_name(
    source: str, module_name: str | None, push: PushFn, host_name: str | None = None
) -> None:
    """A module named a word the VBE refuses. A file can still hold one, and its
    procedures run called bare, but a call through its name does not compile.
    The `Attribute VB_Name` line is marked, or else the first line."""
    lower = module_name.lower() if module_name is not None else ""
    library = lower in _REFERENCED_LIBRARIES or (
        lower in _HOST_LIBRARIES and lower == (host_name if host_name is not None else "Excel").lower()
    )
    if not module_name or (not library and lower not in _REFUSED_MODULE_NAMES):
        return
    attribute = _VB_NAME_ATTRIBUTE_RE.search(source)
    if attribute is not None:
        start = attribute.end() - len(attribute.group(1)) - 2
        end = start + len(attribute.group(1)) + 2
    else:
        start = 0
        line_end = _FIRST_LINE_END_RE.search(source)
        end = max(0, line_end.start() if line_end is not None else len(source))
    push(
        "invalidDeclarationName",
        f"'{module_name}' names an object library every project here references, so it cannot "
        "name a module: the VBE refuses the name (\"Name conflicts with existing module, "
        'project, or object library").'
        if library
        else f"Reserved VBA keyword '{module_name}' cannot name a module: the VBE refuses to add "
        f"one, and a call through the name, {module_name}.Proc, does not compile.",
        Span(start, end),
    )


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

# The object-value branch (propertyLetObjectValue) resolves the final value
# parameter's type through resolve_known_object_assignment_type (host/project
# class resolution) and reports an object-typed Property Let value. The other three
# branches are pure signature/structure checks that need no type surface.


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


def check_property_setter_value_parameters(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    """Property Let/Set setters receive the assigned value through the final
    parameter. A setter with no parameters has no value slot, setters have no return
    type, and Property Set value parameters must be object references. A Property
    Let's value parameter may be of any type (XLIDE issue #107)."""
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
        # A ParamArray cannot be the value: `Property Let P(ParamArray v())` is
        # "Argument not optional" (XLIDE issue #266, measured in Excel 16.0).
        if len(member.params) > 0 and not member.params[-1].param_array:
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
            # A Property Let's value parameter may be any type, object types
            # included: `Property Let Item(ByVal v As Object)`, `As Worksheet` and
            # `As <project class>` all compile, and `h.Item = New Collection` calls
            # the Let (XLIDE issue #107, measured in Excel 16.0). The old
            # property-let-object-value report was wrong and is retired.
            continue
        push(
            "propertySetterMissingValue",
            f"{label} '{member.name}' must take its value in a parameter after its ParamArray. "
            "This is a VBE compile error: Argument not optional."
            if len(member.params) > 0
            else f"{label} '{member.name}' must include a final value parameter.",
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


def _collect_implements_refs(source: str, scan_end: int) -> list[_TypeNameRef]:
    out: list[_TypeNameRef] = []
    line_start = 0
    length = len(source)
    while line_start <= scan_end:
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
    # Implements is only legal in the declarations section, so the line scan
    # stops at the first procedure (typeSemanticTokens.ts). A line inside a
    # procedure that starts with it names no type.
    first_procedure_start = next(
        (member.span.start for member in mod.members if isinstance(member, ProcedureNode)), len(source)
    )
    out.extend(_collect_implements_refs(source, first_procedure_start))

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
    "host": "a host object-model type",
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


# The Scripting Runtime's types, which missing-library-reference judges (XLIDE issue #349).
_SCRIPTING_TYPE_NAMES = frozenset({"dictionary", "filesystemobject", "textstream"})


def check_invalid_as_type_names(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, opts: AnalyzeModuleOptions, push: PushFn) -> None:
    with_events_new_spans = _collect_with_events_new_declaration_spans(mod, activity)
    known_non_type_names = opts.known_non_type_names or frozenset()
    variables: set[str] | None = None
    own_types: set[str] | None = None
    qualified_enum_types: set[str] | None = None
    for ref in _collect_type_name_references(source, mod):
        if activity is not None and activity.is_inactive(ref.span):
            continue
        lookup_name = _type_reference_lookup_name(ref)
        if ref.kind == "declaration" and ref.qualifier:
            if qualified_enum_types is None:
                qualified_enum_types = {f"{type_.module_name}.{type_.name}".lower() for type_ in opts.project_types or () if type_.kind is VbaProjectTypeKind.ENUM and type_.module_name}
                qualified_enum_types.update(f"{opts.module_name or 'Module'}.{member.name}".lower() for member in active_module_members(mod, activity) if isinstance(member, EnumNode))
            if lookup_name.lower() in qualified_enum_types and resolve_type_name(lookup_name, None, opts.host_model) is None:
                push("invalidAsTypeName", f"'{lookup_name}' qualifies a source Enum with its module name. Use '{ref.name}' as the type name. This is a VBE compile error: User-defined type not defined.", ref.span)
                continue
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
        # A Private Type or Enum of another module, bare or qualified, and any
        # name qualified by a variable, are no type here (XLIDE issue #490,
        # measured in Excel 16.0).
        if opts.hidden_type_names is not None and lookup_name.lower() in opts.hidden_type_names:
            push(
                "invalidAsTypeName",
                f"'{lookup_name}' is Private to the module that declares it, so this module cannot "
                "use it as a type. This is a VBE compile error: User-defined type not defined.",
                ref.span,
            )
            continue
        if ref.qualifier:
            if variables is None:
                variables = _declared_variable_names(mod, activity)
            if ref.qualifier.lower() in variables:
                push(
                    "invalidAsTypeName",
                    f"'{ref.qualifier}' is a variable, and a variable never qualifies a type. This "
                    "is a VBE compile error: User-defined type not defined.",
                    ref.span,
                )
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
        # No type of the project and none of a referenced library spells it,
        # where every library the project references is one whose names are
        # all known (XLIDE issue #234, measured in Excel 16.0).
        # The Scripting Runtime's own types are missing-library-reference's, which
        # names the reference to add.
        libraries = (
            [
                cast("AbstractSet[str] | None", library_type_names(library))
                for library in (dict.fromkeys(["VBA", (opts.host_model or {}).get("hostName", opts.host or "Excel"), *opts.referenced_libraries]) if opts.referenced_libraries else opts.referenced_libraries)
            ]
            if opts.referenced_libraries is not None
            else None
        )
        if own_types is None:
            own_types = {
                member.name.lower()
                for member in active_module_members(mod, activity)
                if isinstance(member, (TypeNode, EnumNode))
            }
        name_lower = ref.name.lower()
        if (
            not ref.qualifier
            and name_lower not in _SCRIPTING_TYPE_NAMES
            and name_lower not in own_types
            and libraries is not None
            and len(libraries) > 0
            and all(names is not None and name_lower not in names for names in libraries)
        ):
            push(
                "invalidAsTypeName",
                f"No type of this project and none of the libraries it references is named "
                f"'{ref.name}'. This is a VBE compile error: User-defined type not defined.",
                ref.span,
            )


def _declared_variable_names(mod: ModuleNode, activity: ConditionalActivityTracker | None) -> set[str]:
    """The variables the module declares, at module level or in a procedure, lowercased."""
    out: set[str] = set()

    def add(group: VariableGroupNode) -> None:
        if not group.is_const:
            for decl in group.declarations:
                out.add(decl.name.lower())

    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            add(member)
        elif isinstance(member, ProcedureNode):
            for_each_variable_group(member.body, add, activity)
    return out


# -- checkDimInitializer ---------------------------------------------------


def _top_level_assign_offset(source: str, span: Span) -> int | None:
    # Same token view the rest of the pass uses, so the statement is not
    # re-lexed here (cached_statement_tokens already drops comments/newlines).
    toks = statement_tokens(source, span)
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


@dataclass(frozen=True, slots=True)
class _DeclarationJunk:
    text: str
    span: Span
    why: str
    error: str | None = None


_TYPE_SUFFIX_TOKEN_RE = re.compile(r"^[$%&!#@]$")


def _declaration_junk(source: str, span: Span, is_const: bool) -> _DeclarationJunk | None:
    """What stands after a declared name where the VBE takes nothing: a second
    word, `Dim asdf qwer`; no type name after As, `Private v As 123`; or a
    second value in a Const, `Const K = asdf qwer` (XLIDE issue #234). A complete
    type followed by more is unexpectedTokenAfterDeclarationType's."""
    toks = statement_tokens(source, span)
    i = 1 if token_text(_at(toks, 0)) == "withevents" else 0
    name = _at(toks, i)
    # A name that is no identifier, or runs on into what follows (`_name`,
    # `1value`, `user-name`), is the identifier rules' to report.
    if name is None or not _is_declaration_type_name_token(name):
        return None
    i += 1
    after = _at(toks, i)
    if after is not None and after.start == name.end and _TYPE_SUFFIX_TOKEN_RE.search(after.raw_text):
        i += 1
    elif after is not None and after.start == name.end and after.raw_text != "(":
        return None
    if _raw_at(toks, i) == "(":
        close = match_paren_from(toks, i)
        if close < 0:
            return None
        i = close + 1
    nxt = _at(toks, i)
    if nxt is None:
        return None
    if token_text(nxt) == "as":
        type_tok = _at(toks, i + 2 if token_text(_at(toks, i + 1)) == "new" else i + 1)
        # `As (Long)` is "Syntax error" (XLIDE issue #236).
        if type_tok is not None and not _is_declaration_type_name_token(type_tok):
            return _DeclarationJunk(
                type_tok.raw_text,
                absolute_span(span, type_tok),
                "As needs a type name",
                "Syntax error" if type_tok.raw_text == "(" else "Expected: New or type name",
            )
        return None
    if nxt.raw_text == "=":
        at = juxtaposed_value_index(toks, i + 1) if is_const else -1
        if at < 0:
            return None
        return _DeclarationJunk(
            toks[at].raw_text,
            absolute_span(span, toks[at]),
            "a Const takes one value",
            "Expected: end of statement",
        )
    return _DeclarationJunk(
        nxt.raw_text,
        absolute_span(span, nxt),
        "a declaration takes As and a type there, or nothing",
    )


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

    # A declaration that is not `name [As type]` (XLIDE issue #234, measured in
    # Excel 16.0): "Syntax error" in a procedure, "Expected: end of statement"
    # at module level.
    def inspect_group(group: VariableGroupNode, error: str = "Syntax error") -> None:
        for decl in group.declarations:
            inspect(decl.span, True)
            junk = _declaration_junk(source, decl.span, group.is_const is True)
            if junk is not None:
                push(
                    "unexpectedDeclarationToken",
                    f"Unexpected '{junk.text}' after '{decl.name}': {junk.why}. This is a VBE "
                    f"compile error: {junk.error if junk.error is not None else error}.",
                    junk.span,
                )

    first_procedure = next((m for m in mod.members if isinstance(m, ProcedureNode)), None)
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            # After a procedure the VBE says only "Syntax error".
            inspect_group(
                member,
                "Syntax error"
                if first_procedure is not None and member.span.start > first_procedure.span.start
                else "Expected: end of statement",
            )
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
# string"). The rule resolves decimal integer literal sizes and same
# module/procedure Const aliases that reduce to a decimal integer literal; unknown,
# duplicate, string, and compound constants are deliberately left unresolved (no
# size, so no diagnostic), which keeps the rule free of false positives.
_FIXED_LENGTH_STRING_MIN = 1
_FIXED_LENGTH_STRING_MAX = 65526


def check_fixed_length_string_bounds(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    module_constants = collect_module_literal_integer_constants(mod, activity)

    def inspect_declaration(decl: VariableDeclNode | TypeFieldNode, constants: dict[str, float | None]) -> None:
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
    """An `Option` statement may not follow a procedure.

    Only a procedure closes the window. Measured in Excel 16.0 (build 20326,
    2026-09-26): each of Option Explicit, Base, Compare and Private Module compiles
    after a Const, a module variable, a Type, an Enum, a Declare and a Deftype
    statement (XLIDE issue #113 opened with `DefLng A-Z` above `Option Explicit`),
    and is refused only after `End Sub` / `End Function` with "Only comments may
    appear after End Sub, End Function, or End Property". The rule used to treat
    every declaration as closing the window.
    """
    # Procedures that precede the Option under test AND could be compiled beside it:
    # a procedure in the other arm of a chain closes no window, because only one arm
    # is ever built (XLIDE issue #58).
    procedures_above: list[Span] = []
    # `Option Base` alone has a second closer: a module-level array already
    # dimensioned above it ("Array already dimensioned", measured 2026-09-26).
    arrays_above: list[Span] = []

    def compiled_with(priors: Sequence[Span], member_span: Span) -> bool:
        return any(
            activity is None or not activity.mutually_exclusive(prior, member_span) for prior in priors
        )

    for member in active_module_members(mod, activity):
        if isinstance(member, OptionNode):
            if compiled_with(procedures_above, member.span):
                push(
                    "optionAfterDeclaration",
                    "Option statements must appear before the first procedure; only comments may "
                    "follow End Sub, End Function, or End Property.",
                    first_token_span(source, member.span),
                )
            elif _OPTION_BASE_RE.match(member.option_text.strip()) and compiled_with(arrays_above, member.span):
                push(
                    "optionAfterDeclaration",
                    "'Option Base' must come before any array declaration: an array above it is "
                    "already dimensioned.",
                    first_token_span(source, member.span),
                )
            continue
        if isinstance(member, ProcedureNode):
            procedures_above.append(member.span)
        elif isinstance(member, VariableGroupNode) and any(decl.is_array for decl in member.declarations):
            arrays_above.append(member.span)


_OPTION_BASE_RE = re.compile(r"base\b", re.IGNORECASE)


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
    """A module may declare each Option only once (MS-VBAL 5.2.1; oracle-verified
    `duplicate_option_explicit_compile`). Keyed by category, so two `Option Compare`
    collide even with different arguments. Two Options in different arms of one
    `#If` chain are alternatives, so only the arm matters, not whether the branch
    can be decided: skipping every undecidable branch went blind to a real repeat
    inside one arm."""
    options = [member for member in active_module_members(mod, activity) if isinstance(member, OptionNode)]

    def key_of(member: OptionNode) -> str | None:
        if activity is not None and activity.is_inactive(member.span):
            return None
        return _option_category(member).lower() or None

    def report(repeat: OptionNode, earlier: OptionNode) -> None:
        push(
            "duplicateOption",
            f"Duplicate Option statement; only one 'Option {_option_category(repeat)}' is allowed per module.",
            first_token_span(source, repeat.span),
        )

    report_repeated_keys(options, activity, key_of, lambda member: member.span, report)


def _option_category(member: OptionNode) -> str:
    """The word after `Option`, which is what may appear at most once."""
    parts = member.option_text.strip().split()
    return parts[0] if parts else ""


# -- checkOptionStatementForm ----------------------------------------------


def check_option_statement_form(
    source: str,
    mod: ModuleNode,
    opts: AnalyzeModuleOptions,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """An Option statement names one of the four directives VBA has, takes the
    argument that one takes, and ends there. The parser kept whatever followed
    `Option` without reading it, so `Option Explicit()` analyzed clean while the
    module would not compile (XLIDE issue #74).

    The forms follow the live VBE's compile dialog for each malformed shape (Excel
    oracle probes, 2026-09-13): a bare `Option` or an unknown word expects Base,
    Compare, Explicit or Private; `Option Base` takes 0 or 1; `Option Compare` takes
    Text or Binary; `Option Private` is written `Option Private Module`; and anything
    after a complete form expects the statement to end.

    `Option Compare Database` is the one the host decides. Access writes it into
    every module it creates, and Excel refuses it with the same "Text or Binary" as
    any other unknown argument, so it is reported only where the project names a
    host that is not Access. A file no project claims names no host and is left
    alone.
    """
    host = opts.host.lower() if opts.host else None
    for member in active_module_members(mod, activity):
        if not isinstance(member, OptionNode):
            continue
        # A trailing comment is not trailing junk, and a line continuation is
        # trivia the lexer already attached to the token that follows it.
        toks = statement_tokens(source, member.span)

        def report(index: int, message: str, _toks: list[VbaToken] = toks, _span: Span = member.span) -> None:
            tok = _toks[index] if index < len(_toks) else None
            push(
                "invalidOptionStatement",
                message,
                absolute_span(_span, tok) if tok is not None else first_token_span(source, _span),
            )

        def ends_here(index: int, form: str, _toks: list[VbaToken] = toks) -> None:
            """Report trailing tokens after a directive whose form is complete."""
            if len(_toks) > index:
                report(index, f"'{form}' is complete here; VBA expects the statement to end.")

        def argument(index: int, _toks: list[VbaToken] = toks) -> str | None:
            return token_text(_toks[index]) if index < len(_toks) else None

        # toks[0] is `Option` itself; the directive it names follows it.
        directive = argument(1)
        if directive is None:
            report(0, "'Option' names no directive; VBA expects Base, Compare, Explicit or Private.")
            continue
        if directive == "explicit":
            ends_here(2, "Option Explicit")
        elif directive == "base":
            arg = argument(2)
            if arg not in ("0", "1"):
                report(1 if arg is None else 2, "'Option Base' takes 0 or 1.")
                continue
            ends_here(3, "Option Base")
        elif directive == "compare":
            arg = argument(2)
            if arg == "database":
                if host is not None and host != "access":
                    report(
                        2,
                        "'Option Compare Database' is an Access directive; this project's host "
                        "takes Binary or Text.",
                    )
                    continue
            elif arg not in ("binary", "text"):
                report(1 if arg is None else 2, "'Option Compare' takes Binary or Text.")
                continue
            ends_here(3, "Option Compare")
        elif directive == "private":
            if argument(2) == "module" and is_object_module_kind(opts.module_kind):
                # Measured in Excel 16.0 (XLIDE issue #124): "Option Private Module
                # not permitted in an object module".
                report(2, "'Option Private Module' is not permitted in a class, document or UserForm module.")
                continue
            if argument(2) != "module":
                report(
                    1 if argument(2) is None else 2,
                    "'Option Private' is written 'Option Private Module'.",
                )
                continue
            ends_here(3, "Option Private Module")
        else:
            written = toks[1].canonical_text or toks[1].raw_text
            report(
                1,
                f"'Option {written}' is not an Option statement; "
                "VBA expects Base, Compare, Explicit or Private.",
            )


# -- checkTooManyParameters ------------------------------------------------


def check_too_many_parameters(
    mod: ModuleNode,
    module_kind: ModuleSymbolKind | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Rule: a procedure, Event or Declare may declare at most 60 parameters, and
    at most 59 in a class module. VBE rejects one more with "Too many
    arguments" (oracle-verified `corpus_arg_limit_001b_compile`; XLIDE issue #210,
    measured in Excel 16.0: a class Sub, Friend Sub, Function, Property Get,
    Property Let with its value, Event and Private Declare each compile with
    59 and are refused with 60, and a ParamArray counts as one). Document
    modules and UserForms could not be measured the same way, so they keep 60."""
    limit = (
        _MAX_CLASS_PROCEDURE_PARAMETERS
        if module_kind is ModuleSymbolKind.CLASS
        else _MAX_PROCEDURE_PARAMETERS
    )
    for member in active_module_members(mod, activity):
        if (
            not isinstance(member, (ProcedureNode, EventNode, DeclareNode))
            or len(member.params) <= limit
        ):
            continue
        what = (
            "a procedure"
            if isinstance(member, ProcedureNode)
            else "an Event"
            if isinstance(member, EventNode)
            else "a Declare"
        )
        subject = (
            f"In a class module, {what}"
            if module_kind is ModuleSymbolKind.CLASS
            else f"{what[0].upper()}{what[1:]}"
        )
        push(
            "tooManyParameters",
            f"{subject} may have at most {limit} parameters; "
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
        # A Declare's parameters are held to both (XLIDE issue #253, measured in
        # Excel 16.0: "User-defined type may not be passed ByVal" and "Invalid
        # optional parameter type").
        if not isinstance(member, (ProcedureNode, DeclareNode)):
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
        # The final parameter of a Property Let/Set is the assigned value: mandatory by
        # definition and exempt from the required-after-optional rule (MS-VBAL 5.3.1.5),
        # so an Optional index parameter may legally precede it.
        last_is_value_param = member.proc_kind in (ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET)
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
                # A Property Let or Set takes its value after the ParamArray:
                # `Property Let P(ParamArray v() As Variant, ByVal x As Long)`
                # compiles (XLIDE issue #266, measured in Excel 16.0).
                if i != len(params) - 1 and not (last_is_value_param and i == len(params) - 2):
                    push(
                        "paramArrayNotLast",
                        f"ParamArray '{p.name}' must be the last parameter.",
                        declared_name_span(source, p.span, p.name),
                    )
                continue
            if p.optional:
                optional_seen = True
                continue
            if optional_seen and not (last_is_value_param and i == len(params) - 1):
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
        name_reason = _property_parameter_name_mismatch(expected, actual, i + 1)
        if name_reason:
            return name_reason
    return None


_BRACKET_EDGE_RE = re.compile(r"^\[|\]$")


def _property_parameter_name_mismatch(
    expected: ParameterNode, actual: ParameterNode, index: int
) -> str | None:
    """An index parameter keeps its name across the property's procedures: `Get
    P(ByVal i As Long)` with `Let P(ByVal k As Long, ...)` is "Definitions of
    property procedures for the same property are inconsistent" (XLIDE issue #266,
    measured in Excel 16.0). Case does not count, and the value parameter may
    have any name."""

    def bare(name: str) -> str:
        return _BRACKET_EDGE_RE.sub("", name).lower()

    if bare(expected.name) == bare(actual.name):
        return None
    return (
        f"Index parameter {index} must keep its name: expected '{expected.name}', "
        f"found '{actual.name}'."
    )


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
        if len(group.gets) == 0 and len(group.setters) == 2:
            # A Let and a Set with no Get: their indexes keep one set of names
            # (XLIDE issue #266, measured in Excel 16.0).
            first, second = group.setters
            first_indexes = first.params[:-1]
            second_indexes = second.params[:-1]
            if len(first_indexes) == len(second_indexes):
                for i in range(len(first_indexes)):
                    name_reason = _property_parameter_name_mismatch(
                        first_indexes[i], second_indexes[i], i + 1
                    )
                    if name_reason:
                        push(
                            "propertyAccessorSignatureMismatch",
                            f"{_property_procedure_label(second.proc_kind)} '{second.name}' argument "
                            f"list must match {_property_procedure_label(first.proc_kind)} "
                            f"'{first.name}' before the final value parameter. {name_reason}",
                            declared_name_span(source, second.span, second.name),
                        )
                        break
            continue
        if len(group.gets) != 1:
            continue
        getter = group.gets[0]
        for setter in group.setters:
            if len(setter.params) == 0:
                continue
            reason = _property_index_parameter_mismatch(getter.params, setter.params[:-1])
            if reason is not None:
                push(
                    "propertyAccessorSignatureMismatch",
                    f"{_property_procedure_label(setter.proc_kind)} '{setter.name}' argument list "
                    f"must match Property Get '{getter.name}' before the final value parameter. {reason}",
                    declared_name_span(source, setter.span, setter.name),
                )
                continue
            # A Let's value parameter must have the Get's type: `Get Size() As Long`
            # with `Let Size(ByVal v As Integer)` is "Definitions of property
            # procedures for the same property are inconsistent" (XLIDE issue #124,
            # measured in Excel 16.0). Either side without a type is Variant. A Set's
            # value is never compared: Variant, Object, Collection and even Long Gets
            # beside an Object or Collection Set all compile (XLIDE issue #152).
            if setter.proc_kind is not ProcKind.PROPERTY_LET:
                continue
            value_param = setter.params[-1]
            get_type = normalize_type(getter.return_type)
            if get_type is None and not getter.type_suffix:
                get_type = "variant"
            value_type = normalize_type(value_param.as_type)
            if value_type is None and not value_param.type_suffix:
                value_type = "variant"
            # An object Get beside a Variant Let compiles: `Get M() As Collection`
            # with `Let M(ByVal v As Variant)` (XLIDE issue #414, measured in Excel
            # 16.0). A Collection Get with an Object Let does not.
            object_get_variant_let = (
                value_type == "variant"
                and get_type is not None
                and get_type != "variant"
                and not is_known_scalar_type(get_type)
            )
            if (
                get_type is not None
                and value_type is not None
                and get_type != value_type
                and not value_param.is_array
                and not object_get_variant_let
            ):
                push(
                    "propertyAccessorSignatureMismatch",
                    f"{_property_procedure_label(setter.proc_kind)} '{setter.name}' takes its value As "
                    f"{value_param.as_type if value_param.as_type is not None else 'Variant'}, but Property "
                    f"Get '{getter.name}' returns "
                    f"{getter.return_type if getter.return_type is not None else 'Variant'}; the "
                    "definitions of a property's procedures must agree.",
                    declared_name_span(source, value_param.span, value_param.name),
                )


# -- checkNonConstantConstValues / checkNonConstantEnumMemberValues --------

# Operator keywords (And, Or, Not, Mod, ...) lex as keyword but are never callable
# names; exclude them from the call heuristic so `6 And (3)` is not read as a call.
_OPERATOR_KEYWORD_WORDS: frozenset[str] = frozenset(w.lower() for w in OPERATOR_IDENTIFIERS)


# The intrinsic functions the VBE folds inside an Enum member value and an Optional
# parameter default, where it refuses every call in a Const. Measured one by one in
# Excel 16.0 (build 20326, 2026-09-26; XLIDE issue #112): these sixteen compile in
# both positions, while Asc, AscW, Chr, Val, Sqr, RGB, Round, IIf, Hex, Oct, InStr,
# StrComp, CDec, DateSerial, Choose, Mid, Left, UCase, Str, Trim, Format, Replace,
# String, Space, Now, Timer, Rnd and Array are "Constant expression required" there
# too.
_CONSTANT_FOLDED_INTRINSICS: frozenset[str] = frozenset(
    {
        "len", "lenb", "abs", "int", "fix", "sgn",
        "cint", "clng", "clnglng", "cbyte", "cbool", "cdbl", "csng", "ccur", "cvar", "cdate",
    }
)


def _non_constant_default_element(
    tokens: list[VbaToken], base_offset: int, position: str
) -> tuple[str, Span] | None:
    """The first element of a Const value ("const"), an Enum member value or an
    Optional parameter default ("enumOrOptional") that is not a constant."""
    for i, tok in enumerate(tokens):
        word = (tok.canonical_text if tok.canonical_text is not None else tok.raw_text).lower()
        if tok.kind is TokenKind.KEYWORD and (word == "new" or word == "addressof"):
            return (f"'{tok.raw_text}'", Span(base_offset + tok.start, base_offset + tokens[-1].end))
        is_name = tok.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED_IDENTIFIER)
        is_operator_keyword = tok.kind is TokenKind.KEYWORD and word in _OPERATOR_KEYWORD_WORDS
        if (
            position == "enumOrOptional"
            and word in _CONSTANT_FOLDED_INTRINSICS
            and (i == 0 or tokens[i - 1].raw_text != ".")
        ):
            continue
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        if is_name and not is_operator_keyword and nxt is not None and nxt.raw_text == "(":
            close_index = match_paren_from(tokens, i + 1)
            end_tok = tokens[close_index] if close_index >= 0 else tokens[i + 1]
            return (f"the call '{tok.raw_text}(...)'", Span(base_offset + tok.start, base_offset + end_tok.end))
    return None


def _value_tokens_after_equals(source: str, span: Span) -> tuple[list[VbaToken], Span] | None:
    """The value's tokens after the top-level `=`, and the span they cover."""
    toks = [
        t
        for t in tokenize(source[span.start : span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    eq = top_level_operator_index(toks, "=")
    if eq < 0 or eq + 1 >= len(toks):
        return None
    tokens = toks[eq + 1 :]
    return tokens, span_for_tokens(tokens, span.start)


def check_non_constant_parameter_defaults(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    """An Optional parameter default must be a constant expression (no call/New/AddressOf).

    Object-typed parameters, host and project classes included, are skipped: their
    defaults are owned by the parameter-default-type-mismatch rule ("must be
    Nothing"), which avoids a double diagnostic.
    """
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        for param in member.params:
            if not param.default_raw:
                continue
            if resolve_known_object_assignment_type(param.as_type, member_ctx) is not None:
                continue
            value_tokens = _value_tokens_after_equals(source, param.span)
            if value_tokens is None:
                continue
            non_constant = _non_constant_default_element(value_tokens[0], param.span.start, "enumOrOptional")
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
    """VBA's functions that take no argument, `Now`, `Date`, `Time`, `Timer` and
    `Rnd`, are calls without the parentheses unless the module declares the name
    (XLIDE issue #255, measured)."""
    declared: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode):
            for decl in member.declarations:
                declared.add(decl.name.lower())
        else:
            # `'name' in member && typeof member.name === 'string'`.
            member_name = getattr(member, "name", None)
            if isinstance(member_name, str):
                declared.add(member_name.lower())

    def inspect_group(group: VariableGroupNode) -> None:
        if not group.is_const:
            return
        for decl in group.declarations:
            if decl.default_raw is None or is_inactive_node(activity, decl):
                continue
            value_tokens = _value_tokens_after_equals(source, decl.span)
            if value_tokens is None:
                continue
            tokens = value_tokens[0]
            non_constant = _non_constant_default_element(tokens, decl.span.start, "const")
            if non_constant is None:
                non_constant = _argumentless_function(tokens, decl.span.start, declared)
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


# VBA's functions measured as refused in a Const without parentheses (XLIDE issue #255).
_ARGUMENTLESS_FUNCTIONS = frozenset({"now", "date", "time", "timer", "rnd"})


def _argumentless_function(
    toks: Sequence[VbaToken], base: int, declared: AbstractSet[str]
) -> tuple[str, Span] | None:
    """`Const K = Now`: a VBA function named without parentheses, which still calls it."""
    for i, tok in enumerate(toks):
        word = token_text(tok)
        if word in _ARGUMENTLESS_FUNCTIONS and word not in declared and _raw_at(toks, i - 1) != ".":
            return (
                f"'{tok.raw_text}', a VBA function evaluated as the code runs,",
                Span(base + tok.start, base + tok.end),
            )
    return None


def check_non_constant_enum_member_values(source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn) -> None:
    string_consts: dict[str, str | None] | None = None
    for member in active_module_members(mod, activity):
        if not isinstance(member, EnumNode):
            continue
        if string_consts is None:
            string_consts = _module_string_constants(source, mod, activity)
        for enum_member in member.members:
            if enum_member.value_raw is None or is_inactive_node(activity, enum_member):
                continue
            value_tokens = _value_tokens_after_equals(source, enum_member.span)
            if value_tokens is None:
                continue
            tokens, value_span = value_tokens
            non_constant = _non_constant_default_element(tokens, enum_member.span.start, "enumOrOptional")
            if non_constant is not None:
                label, hit_span = non_constant
                push(
                    "enumMemberNotConstant",
                    f"Enum member '{enum_member.name}' value must be a constant expression; "
                    f"{label} is not constant.",
                    hit_span,
                )
                continue
            text = _constant_string_value(tokens, string_consts)
            if text is not None and _string_is_never_numeric(text):
                push(
                    "enumMemberTypeMismatch",
                    f"Enum member '{enum_member.name}' is the string \"{text}\", and an Enum member "
                    "is a Long. This is a VBE compile error: Type mismatch.",
                    value_span,
                )


def _module_string_constants(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> dict[str, str | None]:
    """The module's Consts whose value is a string: `Const S As String = "x"` or
    one built from others with `&`. Keyed by lowercased name; a name declared
    twice maps to None (undefined)."""
    out: dict[str, str | None] = {}
    for member in active_module_members(mod, activity):
        if not isinstance(member, VariableGroupNode) or not member.is_const:
            continue
        for decl in member.declarations:
            value_tokens = (
                None if decl.default_raw is None else _value_tokens_after_equals(source, decl.span)
            )
            tokens = value_tokens[0] if value_tokens is not None else None
            key = decl.name.lower()
            out[key] = None if key in out or not tokens else _constant_string_value(tokens, out)
    return out


def _constant_string_value(
    tokens: Sequence[VbaToken], string_consts: Mapping[str, str | None] | None
) -> str | None:
    """A constant expression's value when it is a string: a literal, a Const that
    is one, or those joined with `&`. None (undefined) for anything else."""
    parts = [tok for tok in tokens if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE]
    out = ""
    for i, tok in enumerate(parts):
        if i % 2 == 1:
            if tok.raw_text != "&":
                return None
            continue
        if tok.kind is TokenKind.STRING_LITERAL:
            out += tok.raw_text[1:-1].replace('""', '"')
        elif tok.kind is TokenKind.IDENTIFIER:
            value = string_consts.get(tok.raw_text.lower()) if string_consts is not None else None
            if value is None:
                return None
            out += value
        else:
            return None
    return out if len(parts) % 2 == 1 else None


_ASCII_DIGIT_RE = re.compile(r"\d", re.ASCII)


def _string_is_never_numeric(text: str) -> bool:
    """Whether no locale could read the string as a number: it has no digit and is
    not a `&H`/`&O` literal. Measured in Excel 16.0 (XLIDE issue #210): `"1"` and
    `"&H10"` are Longs to VBA, and `"x"`, `""` and `"True"` are a Type mismatch,
    though CLng("True") runs."""
    trimmed = js_trim(text)
    return not _ASCII_DIGIT_RE.search(trimmed) and not trimmed.startswith("&")


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
    # Procedures that precede the declaration under test AND could be compiled
    # beside it. A procedure in one arm of a `#If` chain and a declaration in
    # another arm never reach the compiler together, so the declaration is not
    # "after" it in any build (XLIDE issue #58).
    procedures_above: list[Span] = []
    malformed_conditional_blocks = scan_conditional_compilation_branch_order(mod).malformed_block_spans
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            procedures_above.append(member.span)
            continue
        compiled_together = any(
            activity is None or not activity.mutually_exclusive(prior, member.span)
            for prior in procedures_above
        )
        if not compiled_together:
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
