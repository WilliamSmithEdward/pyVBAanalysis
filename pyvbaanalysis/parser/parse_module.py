"""Error-tolerant VBA parser producing the AST defined in nodes.py.

Ported from xlide_vscode/src/analyzer/parser/parseModule.ts. Verified against
MS-VBAL v20250520 (sections 4.2, 5.2-5.4).

Design notes:
- The parser works over logical statements (MS-VBAL 3.3.1 EOS), so its natural
  recovery points are exactly the newline/colon boundaries. It never throws.
- Block statements track an open-block stack so a stray terminator is reported
  instead of corrupting the tree; an unclosed block yields a "missing End X"
  diagnostic while still returning a node.
- The structured-vs-raw fallback (agent.md Risk 4) is matched exactly: a body
  statement becomes a structured Assignment/Call only when it parses cleanly and
  fully; anything else stays a raw StatementNode.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass, field

from ..lexer.token_helpers import (
    match_paren_from,
    split_top_level_token_groups,
    token_word,
    tokens_without_leading_line_number,
)
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize
from .fixed_length_string import parse_fixed_length_string_type
from .nodes import (
    AssignmentNode,
    AttributeNode,
    BodyNode,
    CallNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    DeclareNode,
    DoBlockNode,
    EnumMemberNode,
    EnumNode,
    EventNode,
    ExprNode,
    ForBlockNode,
    IfBlockNode,
    IfBranchKind,
    IfBranchNode,
    IndexExpr,
    ModuleKind,
    ModuleMember,
    ModuleNode,
    OptionNode,
    ParameterNode,
    ParseDiagnostic,
    ParseSeverity,
    ProcedureNode,
    ProcKind,
    SelectBlockNode,
    Span,
    StatementNode,
    TypeFieldNode,
    TypeNode,
    VariableDeclNode,
    VariableGroupNode,
    WhileBlockNode,
    WithBlockNode,
)
from .parse_expression import ExprParseResult, parse_expression, parse_parenless_arguments
from .parser_state import LogicalStatement, StatementCursor, code_tokens, split_logical_statements
from .type_declaration_suffix import is_type_declaration_suffix, type_name_for_declaration_suffix


@dataclass(slots=True)
class _DeclaredNameInfo:
    name: str
    next_index: int
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None


@dataclass(slots=True)
class _IfBranchBuilder:
    """Mutable accumulator for one If-block arm while its body is being parsed."""

    branch_kind: IfBranchKind
    condition: ExprNode | None
    header_span: Span
    body: list[BodyNode] = field(default_factory=list)
    condition_raw: str | None = None
    condition_span: Span | None = None


# Visibility / sharing modifiers that may lead a declaration (MS-VBAL 5.2.3).
_LEADING_MODIFIERS: frozenset[str] = frozenset({"public", "private", "friend", "global", "static"})

# Parameter mechanism markers (MS-VBAL 5.3.1.x).
_PARAM_MARKERS: frozenset[str] = frozenset({"optional", "byval", "byref", "paramarray"})

# Maps an expected-closer tag to its canonical label for diagnostics.
_CLOSER_LABELS: dict[str, str] = {
    "endif": "End If",
    "next": "Next",
    "loop": "Loop",
    "wend": "Wend",
    "endwith": "End With",
    "endselect": "End Select",
    "endsub": "End Sub",
    "endfunction": "End Function",
    "endproperty": "End Property",
    "endtype": "End Type",
    "endenum": "End Enum",
}

_BLOCK_CLOSERS: dict[str, str] = {
    "if": "endif",
    "for": "next",
    "foreach": "next",
    "do": "loop",
    "while": "wend",
    "with": "endwith",
    "select": "endselect",
}

_END_CLOSERS: dict[str, str] = {
    "if": "endif",
    "with": "endwith",
    "select": "endselect",
    "sub": "endsub",
    "function": "endfunction",
    "property": "endproperty",
    "type": "endtype",
    "enum": "endenum",
}

_VB_MEMBER_ATTR_RE = re.compile(r"^VB_[A-Za-z0-9_]+$", re.IGNORECASE)
_VB_CLASS_ATTR_RE = re.compile(r"^VB_(Exposed|Creatable|PredeclaredId)$", re.IGNORECASE)


def _at(tokens: Sequence[VbaToken], i: int) -> VbaToken | None:
    """Safe token access: returns None where TS indexing yields undefined."""
    return tokens[i] if 0 <= i < len(tokens) else None


# The editor surfaces re-parse the same module text many times per request, so a
# small value-keyed LRU collapses those parses to one. The AST is treated as
# immutable by all consumers; callers must not mutate the returned nodes.
_PARSE_CACHE_MAX = 8
_parse_cache: list[tuple[str, ModuleNode]] = []


def parse_module(source: str) -> ModuleNode:
    """Parse VBA source text into a ModuleNode AST. Never throws."""
    for i, (cached_source, _module) in enumerate(_parse_cache):
        if cached_source == source:
            if i > 0:
                _parse_cache.insert(0, _parse_cache.pop(i))
            return _parse_cache[0][1]
    module = _Parser(source, tokenize(source)).parse()
    _parse_cache.insert(0, (source, module))
    if len(_parse_cache) > _PARSE_CACHE_MAX:
        _parse_cache.pop()
    return module


class _Parser:
    _source: str
    _cursor: StatementCursor
    _diagnostics: list[ParseDiagnostic]
    # Expected closers of the currently open blocks (innermost last).
    _open_stack: list[str]

    __slots__ = ("_source", "_cursor", "_diagnostics", "_open_stack")

    def __init__(self, source: str, tokens: Sequence[VbaToken]) -> None:
        self._source = source
        self._cursor = StatementCursor(split_logical_statements(tokens))
        self._diagnostics = []
        self._open_stack = []

    def parse(self) -> ModuleNode:
        members: list[ModuleMember] = []
        while not self._cursor.at_end():
            member = self._parse_module_member()
            if member is not None:
                members.append(member)
        return ModuleNode(
            span=Span(0, len(self._source)),
            module_kind=self._detect_module_kind(members),
            members=members,
            diagnostics=self._diagnostics,
        )

    # -- Module level ------------------------------------------------------

    def _parse_module_member(self) -> ModuleMember | None:
        stmt = self._cursor.peek()
        if stmt is None:
            return None
        tokens = code_tokens(stmt)
        if len(tokens) == 0:
            self._cursor.next()
            return None

        if self._is_attribute(tokens):
            self._cursor.next()
            return self._parse_attribute(stmt, tokens)
        if self._is_conditional_directive(tokens):
            self._cursor.next()
            return self._parse_conditional_directive(stmt, tokens)

        mod_index = self._leading_modifier_count(tokens)
        head = token_word(_at(tokens, mod_index))

        if head == "option":
            self._cursor.next()
            return self._parse_option(stmt, tokens)
        if head == "declare":
            self._cursor.next()
            return self._parse_declare(stmt, tokens, mod_index)
        if head == "event":
            self._cursor.next()
            return self._parse_event(stmt, tokens, mod_index)
        if head == "type":
            return self._parse_type_block(mod_index)
        if head == "enum":
            return self._parse_enum_block(mod_index)
        if head in ("sub", "function", "property"):
            return self._parse_procedure()
        if head == "const":
            self._cursor.next()
            return self._parse_variable_group(stmt, tokens, mod_index, True)
        if head == "dim":
            self._cursor.next()
            return self._parse_variable_group(stmt, tokens, mod_index, False)
        # "Public Foo As Long": a module variable with only a visibility modifier
        # and no Dim keyword (MS-VBAL 5.2.3).
        if mod_index > 0 and _at(tokens, mod_index) is not None:
            self._cursor.next()
            return self._parse_variable_group(stmt, tokens, mod_index, False)
        self._cursor.next()
        return self._make_statement(stmt)

    @staticmethod
    def _is_attribute(tokens: Sequence[VbaToken]) -> bool:
        return len(tokens) >= 1 and token_word(tokens[0]) == "attribute"

    def _parse_attribute(self, stmt: LogicalStatement, tokens: Sequence[VbaToken]) -> AttributeNode:
        eq_index = next((i for i, t in enumerate(tokens) if t.raw_text == "="), -1)
        name_start = _at(tokens, 1)
        name_end_index = eq_index - 1 if eq_index > 1 else len(tokens) - 1
        name_end = _at(tokens, name_end_index) if name_end_index >= 1 else None
        if name_start is not None and name_end is not None:
            upper = eq_index if eq_index > 1 else len(tokens)
            name = "".join(t.raw_text for t in tokens[1:upper])
        else:
            name = ""
        if eq_index >= 0 and eq_index + 1 < len(tokens):
            value_raw = self._source[tokens[eq_index + 1].start : tokens[len(tokens) - 1].end]
        else:
            value_raw = ""
        if name_start is not None and name_end is not None:
            name_span = Span(name_start.start, name_end.end)
        else:
            name_span = Span(stmt.start, stmt.start)
        return AttributeNode(
            span=Span(stmt.start, stmt.end), name=name, name_span=name_span, value_raw=value_raw
        )

    def _parse_option(self, stmt: LogicalStatement, tokens: Sequence[VbaToken]) -> OptionNode:
        option_text = " ".join(
            (t.canonical_text if t.canonical_text is not None else t.raw_text) for t in tokens[1:]
        )
        return OptionNode(span=Span(stmt.start, stmt.end), option_text=option_text)

    def _parse_declare(
        self, stmt: LogicalStatement, tokens: Sequence[VbaToken], mod_index: int
    ) -> DeclareNode:
        visibility = self._canonical(_at(tokens, 0)) if mod_index > 0 else None
        kind_index = -1
        for i, t in enumerate(tokens):
            w = token_word(t)
            if w == "sub" or w == "function":
                kind_index = i
                break
        is_function = kind_index >= 0 and token_word(tokens[kind_index]) == "function"
        name_token = _at(tokens, kind_index + 1) if kind_index >= 0 else None
        name_index = kind_index + 1 if name_token is not None else -1
        if is_function and name_index >= 0:
            declared_name = self._parse_declared_name(tokens, name_index, True)
        else:
            declared_name = _DeclaredNameInfo(
                name=self._strip_brackets(name_token.raw_text) if name_token is not None else "",
                name_span=self._token_span(name_token),
                next_index=name_index + 1,
            )
        ptr_safe = kind_index >= 0 and any(
            token_word(t) == "ptrsafe" for t in tokens[mod_index + 1 : kind_index]
        )
        lib_index = self._find_word_from(tokens, "lib", declared_name.next_index) if name_index >= 0 else -1
        alias_index = (
            self._find_word_from(tokens, "alias", declared_name.next_index) if name_index >= 0 else -1
        )
        lib_name = self._string_literal_text(_at(tokens, lib_index + 1)) if lib_index >= 0 else None
        alias_name = self._string_literal_text(_at(tokens, alias_index + 1)) if alias_index >= 0 else None

        params: list[ParameterNode] = []
        after_paren = declared_name.next_index
        lparen = self._find_raw_after(tokens, "(", name_index) if name_index >= 0 else -1
        if lparen >= 0:
            params, close_index = self._parse_param_list(tokens, lparen)
            after_paren = close_index + 1

        has_as_clause = False
        ap_tok = _at(tokens, after_paren)
        if is_function and ap_tok is not None and token_word(ap_tok) == "as":
            has_as_clause = True
            return_type = self._capture_type(tokens, after_paren + 1)
        else:
            return_type = type_name_for_declaration_suffix(declared_name.type_suffix)
        return DeclareNode(
            span=Span(stmt.start, stmt.end),
            name=declared_name.name,
            is_function=is_function,
            ptr_safe=ptr_safe,
            name_span=declared_name.name_span,
            type_suffix=declared_name.type_suffix,
            type_suffix_span=declared_name.type_suffix_span,
            has_as_clause=has_as_clause,
            visibility=visibility,
            lib_name=lib_name,
            alias_name=alias_name,
            params=params,
            return_type=return_type,
        )

    def _parse_event(
        self, stmt: LogicalStatement, tokens: Sequence[VbaToken], mod_index: int
    ) -> EventNode:
        visibility = self._canonical(_at(tokens, 0)) if mod_index > 0 else None
        name_token = _at(tokens, mod_index + 1)
        name = self._strip_brackets(name_token.raw_text) if name_token is not None else ""
        params: list[ParameterNode] = []
        lparen = self._find_raw_after(tokens, "(", mod_index + 1) if name_token is not None else -1
        if lparen >= 0:
            params, _close = self._parse_param_list(tokens, lparen)
        return EventNode(
            span=Span(stmt.start, stmt.end),
            name=name,
            name_span=self._token_span(name_token),
            visibility=visibility,
            params=params,
        )

    @staticmethod
    def _is_conditional_directive(tokens: Sequence[VbaToken]) -> bool:
        first = _at(tokens, 0)
        return first is not None and first.kind is TokenKind.DIRECTIVE

    def _parse_conditional_directive(
        self, stmt: LogicalStatement, tokens: Sequence[VbaToken]
    ) -> ConditionalDirectiveNode:
        directive_word = token_word(_at(tokens, 1))
        span = Span(stmt.start, stmt.end)
        if directive_word == "const":
            name_token = _at(tokens, 2)
            eq_index = self._find_raw_after(tokens, "=", 2)
            if eq_index >= 0:
                raw, value_span = self._token_range_raw(tokens, eq_index + 1, len(tokens))
            else:
                raw, value_span = (None, None)
            return ConditionalDirectiveNode(
                span=span,
                directive_kind=ConditionalDirectiveKind.CONST,
                name=self._strip_brackets(name_token.raw_text) if name_token is not None else None,
                name_span=Span(name_token.start, name_token.end) if name_token is not None else None,
                value_raw=raw,
                value_span=value_span,
            )
        if directive_word == "if":
            raw, cond_span = self._directive_condition(tokens, 2)
            return ConditionalDirectiveNode(
                span=span, directive_kind=ConditionalDirectiveKind.IF, condition_raw=raw, condition_span=cond_span
            )
        if directive_word == "elseif":
            raw, cond_span = self._directive_condition(tokens, 2)
            return ConditionalDirectiveNode(
                span=span, directive_kind=ConditionalDirectiveKind.ELSE_IF, condition_raw=raw, condition_span=cond_span
            )
        if directive_word == "else":
            return ConditionalDirectiveNode(span=span, directive_kind=ConditionalDirectiveKind.ELSE)
        if directive_word == "end":
            if token_word(_at(tokens, 2)) == "if":
                return ConditionalDirectiveNode(span=span, directive_kind=ConditionalDirectiveKind.END_IF)
        elif directive_word == "endif":
            return ConditionalDirectiveNode(span=span, directive_kind=ConditionalDirectiveKind.END_IF)
        raw, unknown_span = self._token_range_raw(tokens, 1, len(tokens))
        return ConditionalDirectiveNode(
            span=span, directive_kind=ConditionalDirectiveKind.UNKNOWN, condition_raw=raw, condition_span=unknown_span
        )

    def _close_dangling_block_directives(
        self, directives: list[ConditionalDirectiveNode], boundary_offset: int
    ) -> None:
        depth = 0
        for directive in directives:
            if directive.directive_kind is ConditionalDirectiveKind.IF:
                depth += 1
            elif directive.directive_kind is ConditionalDirectiveKind.END_IF:
                depth = max(0, depth - 1)
        for _ in range(depth):
            directives.append(
                ConditionalDirectiveNode(
                    span=Span(boundary_offset, boundary_offset),
                    directive_kind=ConditionalDirectiveKind.END_IF,
                )
            )

    def _parse_variable_group(
        self, stmt: LogicalStatement, tokens: Sequence[VbaToken], mod_index: int, is_const: bool
    ) -> VariableGroupNode:
        i = mod_index
        head = token_word(_at(tokens, i))
        if mod_index > 0:
            modifier = self._canonical(_at(tokens, 0))
        elif head == "dim":
            modifier = "Dim"
        else:
            modifier = ""
        if head == "dim" or head == "const":
            i += 1
        with_events = False
        if token_word(_at(tokens, i)) == "withevents":
            with_events = True
            i += 1
        declarations = self._parse_declarator_list(tokens, i, is_const)
        return VariableGroupNode(
            span=Span(stmt.start, stmt.end),
            modifier=modifier,
            is_const=is_const,
            with_events=with_events,
            declarations=declarations,
        )

    def _parse_declarator_list(
        self, tokens: Sequence[VbaToken], from_: int, is_const: bool
    ) -> list[VariableDeclNode]:
        groups = self._split_top_level_commas(tokens, from_, len(tokens))
        declarations: list[VariableDeclNode] = []
        for group in groups:
            if len(group) == 0:
                continue
            declarations.append(self._parse_declarator(group, is_const))
        return declarations

    def _parse_declarator(self, group: Sequence[VbaToken], is_const: bool) -> VariableDeclNode:
        declared_name = self._parse_declared_name(group, 0, True)
        name = declared_name.name
        i = declared_name.next_index
        is_array = False
        array_bounds: str | None = None
        gi = _at(group, i)
        if gi is not None and gi.raw_text == "(":
            is_array = True
            close = self._match_paren(group, i)
            gclose = _at(group, close)
            if gclose is not None and gclose.raw_text == ")":
                bounds = self._source[group[i].end : group[close].start].strip()
                array_bounds = bounds or None
            i = close + 1
        is_new = False
        as_type: str | None = None
        fixed_length: str | None = None
        has_as_clause = False
        gi = _at(group, i)
        if gi is not None and token_word(gi) == "as":
            has_as_clause = True
            i += 1
            gi2 = _at(group, i)
            if gi2 is not None and token_word(gi2) == "new":
                is_new = True
                i += 1
            as_type, fixed_length = self._capture_declaration_type(group, i, "=" if is_const else None)
        else:
            as_type = type_name_for_declaration_suffix(declared_name.type_suffix)
        default_raw: str | None = None
        eq = next((j for j, t in enumerate(group) if t.raw_text == "="), -1)
        if eq >= 0 and eq + 1 < len(group):
            default_raw = self._source[group[eq + 1].start : group[len(group) - 1].end]
        first = group[0]
        last = group[len(group) - 1]
        return VariableDeclNode(
            span=Span(first.start, last.end),
            name=name,
            name_span=declared_name.name_span,
            type_suffix=declared_name.type_suffix,
            type_suffix_span=declared_name.type_suffix_span,
            has_as_clause=has_as_clause,
            as_type=as_type,
            fixed_length=fixed_length,
            default_raw=default_raw,
            is_array=is_array,
            array_bounds=array_bounds,
            is_new=is_new,
        )

    # -- Type / Enum -------------------------------------------------------

    def _parse_type_block(self, mod_index: int) -> TypeNode:
        head = self._cursor.next()
        assert head is not None
        tokens = code_tokens(head)
        visibility = self._canonical(_at(tokens, 0)) if mod_index > 0 else None
        name_token = _at(tokens, mod_index + 1)
        name = self._strip_brackets(name_token.raw_text) if name_token is not None else ""
        fields: list[TypeFieldNode] = []
        directives: list[ConditionalDirectiveNode] = []
        closed = False
        end_stmt: LogicalStatement | None = None
        while not self._cursor.at_end():
            stmt = self._cursor.peek()
            assert stmt is not None
            if self._closer_kind(stmt) == "endtype":
                end_stmt = self._cursor.next()
                closed = True
                break
            if self._is_module_level_starter(stmt) and not self._is_type_field_statement(stmt):
                break
            self._cursor.next()
            ftokens = code_tokens(stmt)
            if len(ftokens) > 0:
                if self._is_conditional_directive(ftokens):
                    directives.append(self._parse_conditional_directive(stmt, ftokens))
                    continue
                fields.append(self._parse_type_field(stmt, ftokens))
        if not closed:
            self._diag(head, "Type block is missing End Type.", ParseSeverity.ERROR, "MS-VBAL 5.2.3.3")
        elif end_stmt is not None:
            self._close_dangling_block_directives(directives, end_stmt.start)
        end_off = (end_stmt if end_stmt is not None else head).end
        return TypeNode(
            span=Span(head.start, end_off),
            name=name,
            closed=closed,
            name_span=self._token_span(name_token),
            visibility=visibility,
            fields=fields,
            directives=directives if len(directives) > 0 else None,
        )

    def _parse_type_field(self, stmt: LogicalStatement, tokens: Sequence[VbaToken]) -> TypeFieldNode:
        declared_name = self._parse_declared_name(tokens, 0, True)
        name = declared_name.name
        i = declared_name.next_index
        is_array = False
        ti = _at(tokens, i)
        if ti is not None and ti.raw_text == "(":
            is_array = True
            i = self._skip_parens(tokens, i)
        as_type: str | None = None
        fixed_length: str | None = None
        has_as_clause = False
        ti = _at(tokens, i)
        if ti is not None and token_word(ti) == "as":
            has_as_clause = True
            as_type, fixed_length = self._capture_declaration_type(tokens, i + 1)
        else:
            as_type = type_name_for_declaration_suffix(declared_name.type_suffix)
        return TypeFieldNode(
            span=Span(stmt.start, stmt.end),
            name=name,
            name_span=declared_name.name_span,
            type_suffix=declared_name.type_suffix,
            type_suffix_span=declared_name.type_suffix_span,
            has_as_clause=has_as_clause,
            as_type=as_type,
            fixed_length=fixed_length,
            is_array=is_array,
        )

    def _is_type_field_statement(self, stmt: LogicalStatement) -> bool:
        tokens = code_tokens(stmt)
        if len(tokens) == 0:
            return False
        first = token_word(tokens[0])
        if first == "type" and token_word(_at(tokens, 1)) == "as":
            return True
        return any(index > 0 and token_word(token) == "as" for index, token in enumerate(tokens))

    def _parse_enum_block(self, mod_index: int) -> EnumNode:
        head = self._cursor.next()
        assert head is not None
        tokens = code_tokens(head)
        visibility = self._canonical(_at(tokens, 0)) if mod_index > 0 else None
        name_token = _at(tokens, mod_index + 1)
        name = self._strip_brackets(name_token.raw_text) if name_token is not None else ""
        members: list[EnumMemberNode] = []
        directives: list[ConditionalDirectiveNode] = []
        closed = False
        end_stmt: LogicalStatement | None = None
        while not self._cursor.at_end():
            stmt = self._cursor.peek()
            assert stmt is not None
            if self._closer_kind(stmt) == "endenum":
                end_stmt = self._cursor.next()
                closed = True
                break
            if self._is_module_level_starter(stmt):
                break
            self._cursor.next()
            mtokens = code_tokens(stmt)
            if len(mtokens) > 0:
                if self._is_conditional_directive(mtokens):
                    directives.append(self._parse_conditional_directive(stmt, mtokens))
                    continue
                eq_index = next((j for j, t in enumerate(mtokens) if t.raw_text == "="), -1)
                if eq_index >= 0 and eq_index + 1 < len(mtokens):
                    value_raw: str | None = self._source[
                        mtokens[eq_index + 1].start : mtokens[len(mtokens) - 1].end
                    ]
                else:
                    value_raw = None
                members.append(
                    EnumMemberNode(
                        span=Span(stmt.start, stmt.end),
                        name=self._strip_brackets(mtokens[0].raw_text),
                        name_span=self._token_span(mtokens[0]),
                        value_raw=value_raw,
                    )
                )
        if not closed:
            self._diag(head, "Enum block is missing End Enum.", ParseSeverity.ERROR, "MS-VBAL 5.2.3.4")
        elif end_stmt is not None:
            self._close_dangling_block_directives(directives, end_stmt.start)
        end_off = (end_stmt if end_stmt is not None else head).end
        return EnumNode(
            span=Span(head.start, end_off),
            name=name,
            closed=closed,
            name_span=self._token_span(name_token),
            visibility=visibility,
            members=members,
            directives=directives if len(directives) > 0 else None,
        )

    # -- Procedures --------------------------------------------------------

    def _parse_procedure(self) -> ProcedureNode:
        head = self._cursor.next()
        assert head is not None
        tokens = code_tokens(head)
        mod_index = self._leading_modifier_count(tokens)
        modifiers = [self._canonical(t) for t in tokens[0:mod_index]]

        i = mod_index
        head_word = token_word(_at(tokens, i))
        proc_kind: ProcKind
        if head_word == "property":
            i += 1
            accessor = token_word(_at(tokens, i))
            if accessor == "get":
                proc_kind = ProcKind.PROPERTY_GET
            elif accessor == "set":
                proc_kind = ProcKind.PROPERTY_SET
            else:
                proc_kind = ProcKind.PROPERTY_LET
            i += 1
        elif head_word == "function":
            proc_kind = ProcKind.FUNCTION
            i += 1
        else:
            proc_kind = ProcKind.SUB
            i += 1

        declared_name = self._parse_declared_name(
            tokens, i, proc_kind is ProcKind.FUNCTION or proc_kind is ProcKind.PROPERTY_GET
        )
        name = declared_name.name
        i = declared_name.next_index

        params: list[ParameterNode] = []
        after_paren = i
        ti = _at(tokens, i)
        if ti is not None and ti.raw_text == "(":
            params, close_index = self._parse_param_list(tokens, i)
            after_paren = close_index + 1

        has_as_clause = False
        ap = _at(tokens, after_paren)
        if ap is not None and token_word(ap) == "as":
            has_as_clause = True
            return_type = self._capture_type(tokens, after_paren + 1)
        else:
            return_type = type_name_for_declaration_suffix(declared_name.type_suffix)

        expected = self._proc_closer(proc_kind)
        self._open_stack.append(expected)
        body: list[BodyNode] = []
        attributes: list[AttributeNode] = []
        closed = False
        end_stmt: LogicalStatement | None = None
        saw_conditional_directive = False
        while not self._cursor.at_end():
            stmt = self._cursor.peek()
            assert stmt is not None
            ck = self._closer_kind(stmt)
            if ck == expected:
                end_stmt = self._cursor.next()
                closed = True
                break
            stmt_tokens = code_tokens(stmt)
            if self._is_attribute(stmt_tokens):
                if self._is_exported_procedure_attribute(stmt, stmt_tokens, name, len(body) == 0):
                    self._cursor.next()
                    attributes.append(self._parse_attribute(stmt, stmt_tokens))
                    continue
                item = self._parse_body_item(stmt)
                if item is None:
                    break
                body.append(item)
                continue
            nested = self._nested_type_or_enum_block_kind(stmt)
            if nested is not None:
                body.append(self._parse_invalid_nested_module_block_statement(nested, [expected]))
                continue
            # Recovery: a new module-level construct means the End was forgotten.
            if self._is_module_level_starter(stmt):
                if saw_conditional_directive and self._is_alternative_procedure_header(stmt, proc_kind, name):
                    self._cursor.next()
                    continue
                break
            item = self._parse_body_item(stmt)
            if item is None:
                break
            if isinstance(item, ConditionalDirectiveNode):
                saw_conditional_directive = True
            body.append(item)
        self._open_stack.pop()
        if not closed:
            self._diag(
                head,
                f"Procedure '{name}' is missing {_CLOSER_LABELS[expected]}.",
                ParseSeverity.ERROR,
                "MS-VBAL 5.3.1",
            )
        last_body = body[len(body) - 1] if body else None
        if end_stmt is not None:
            end = end_stmt.end
        elif last_body is not None:
            end = last_body.span.end
        else:
            end = head.end
        return ProcedureNode(
            span=Span(head.start, end),
            proc_kind=proc_kind,
            name=name,
            closed=closed,
            name_span=declared_name.name_span,
            type_suffix=declared_name.type_suffix,
            type_suffix_span=declared_name.type_suffix_span,
            has_as_clause=has_as_clause,
            modifiers=modifiers,
            params=params,
            return_type=return_type,
            attributes=attributes if len(attributes) > 0 else None,
            body=body,
        )

    @staticmethod
    def _proc_closer(kind: ProcKind) -> str:
        if kind is ProcKind.FUNCTION:
            return "endfunction"
        if kind is ProcKind.SUB:
            return "endsub"
        return "endproperty"

    def _is_alternative_procedure_header(
        self, stmt: LogicalStatement, current_kind: ProcKind, current_name: str
    ) -> bool:
        tokens = code_tokens(stmt)
        mod_index = self._leading_modifier_count(tokens)
        i = mod_index
        head_word = token_word(_at(tokens, i))
        kind: ProcKind | None = None
        if head_word == "property":
            i += 1
            accessor = token_word(_at(tokens, i))
            if accessor == "get":
                kind = ProcKind.PROPERTY_GET
            elif accessor == "set":
                kind = ProcKind.PROPERTY_SET
            elif accessor == "let":
                kind = ProcKind.PROPERTY_LET
            else:
                kind = None
            i += 1
        elif head_word == "function":
            kind = ProcKind.FUNCTION
            i += 1
        elif head_word == "sub":
            kind = ProcKind.SUB
            i += 1
        name_token = _at(tokens, i)
        name = self._strip_brackets(name_token.raw_text) if name_token is not None else ""
        return kind == current_kind and name.lower() == current_name.lower()

    def _parse_param_list(
        self, tokens: Sequence[VbaToken], lparen: int
    ) -> tuple[list[ParameterNode], int]:
        close_index = self._match_paren(tokens, lparen)
        groups = self._split_top_level_commas(tokens, lparen + 1, close_index)
        params = [self._parse_param(g) for g in groups if len(g) > 0]
        return params, close_index

    def _parse_param(self, group: Sequence[VbaToken]) -> ParameterNode:
        i = 0
        optional = by_val = by_ref = param_array = False
        while True:
            gi = _at(group, i)
            if gi is None or token_word(gi) not in _PARAM_MARKERS:
                break
            w = token_word(gi)
            if w == "optional":
                optional = True
            elif w == "byval":
                by_val = True
            elif w == "byref":
                by_ref = True
            elif w == "paramarray":
                param_array = True
            i += 1
        declared_name = self._parse_declared_name(group, i, True)
        name = declared_name.name
        i = declared_name.next_index
        is_array = False
        gi = _at(group, i)
        if gi is not None and gi.raw_text == "(":
            is_array = True
            i = self._skip_parens(group, i)
        as_type: str | None = None
        has_as_clause = False
        gi = _at(group, i)
        if gi is not None and token_word(gi) == "as":
            has_as_clause = True
            i += 1
            as_type = self._capture_type(group, i, "=")
        else:
            as_type = type_name_for_declaration_suffix(declared_name.type_suffix)
        default_raw: str | None = None
        eq = next((j for j, t in enumerate(group) if t.raw_text == "="), -1)
        if eq >= 0 and eq + 1 < len(group):
            default_raw = self._source[group[eq + 1].start : group[len(group) - 1].end]
        first = group[0]
        last = group[len(group) - 1]
        return ParameterNode(
            span=Span(first.start, last.end),
            name=name,
            name_span=declared_name.name_span,
            type_suffix=declared_name.type_suffix,
            type_suffix_span=declared_name.type_suffix_span,
            has_as_clause=has_as_clause,
            optional=optional,
            by_val=by_val,
            by_ref=by_ref,
            param_array=param_array,
            as_type=as_type,
            is_array=is_array,
            default_raw=default_raw,
        )

    # -- Procedure body items / block statements ---------------------------

    def _parse_body_item(self, stmt: LogicalStatement) -> BodyNode | None:
        ck = self._closer_kind(stmt)
        if ck:
            if ck in self._open_stack:
                # Belongs to an ancestor block; stop and let it close.
                return None
            self._diag(
                stmt,
                f"Unexpected '{_CLOSER_LABELS[ck]}' without a matching opening block.",
                ParseSeverity.ERROR,
                "MS-VBAL 5.4",
            )
            self._cursor.next()
            return self._make_statement(stmt)
        opener = self._opener_kind(stmt)
        if opener is not None:
            return self._parse_block(opener)
        tokens = _code_tokens_after_line_number(stmt)
        if self._is_conditional_directive(tokens):
            self._cursor.next()
            return self._parse_conditional_directive(stmt, tokens)
        head = token_word(_at(tokens, 0))
        if head == "dim" or head == "const" or head == "static":
            self._cursor.next()
            mod_index = 1 if head == "static" else 0
            return self._parse_variable_group(stmt, tokens, mod_index, head == "const")
        structured = self._parse_assignment_or_call(stmt, tokens)
        if structured is not None:
            self._cursor.next()
            return structured
        self._cursor.next()
        return self._make_statement(stmt)

    def _parse_assignment_or_call(
        self, stmt: LogicalStatement, tokens: Sequence[VbaToken]
    ) -> AssignmentNode | CallNode | None:
        if len(tokens) == 0:
            return None
        span = Span(stmt.start, stmt.end)

        # Optional leading Set / Let selects the assignment form.
        lhs_start = 0
        is_set = False
        is_let = False
        head = token_word(tokens[0])
        if head == "set":
            is_set = True
            lhs_start = 1
        elif head == "let":
            is_let = True
            lhs_start = 1

        # Assignment first: any top-level '=' is the assignment operator.
        eq_index = _top_level_equals_index(tokens, lhs_start)
        if eq_index >= 0:
            # Mid(...) = / Mid$(...) = is a dedicated statement form, not assignment.
            if _is_mid_statement_target(tokens, lhs_start):
                return None
            lhs = parse_expression(tokens, lhs_start, eq_index)
            if not _fully_consumed(lhs, eq_index):
                return None
            rhs = parse_expression(tokens, eq_index + 1, len(tokens))
            if not _fully_consumed(rhs, len(tokens)):
                return None
            assert lhs.expr is not None and rhs.expr is not None
            return AssignmentNode(span=span, is_set=is_set, is_let=is_let, lhs=lhs.expr, rhs=rhs.expr)

        # A leading Set/Let with no '=' is a malformed assignment - leave it raw.
        if lhs_start != 0:
            return None
        return self._parse_call_statement(tokens, span)

    def _parse_call_statement(self, tokens: Sequence[VbaToken], span: Span) -> CallNode | None:
        callee_start = 0
        has_call_keyword = False
        if token_word(tokens[0]) == "call":
            has_call_keyword = True
            callee_start = 1
        parsed = parse_expression(tokens, callee_start, len(tokens))
        if parsed.expr is None or len(parsed.diagnostics) > 0:
            return None

        if parsed.end_index == len(tokens):
            # The whole remainder is one expression.
            if isinstance(parsed.expr, IndexExpr):
                return CallNode(
                    span=span,
                    has_call_keyword=has_call_keyword,
                    callee=parsed.expr.callee,
                    args=parsed.expr.args,
                )
            if has_call_keyword:
                return CallNode(span=span, has_call_keyword=True, callee=parsed.expr, args=[])
            # A bare identifier / member chain with no args is ambiguous - keep raw.
            return None

        # Trailing tokens after the callee: parenless argument list (implicit only).
        if has_call_keyword:
            return None
        args = parse_parenless_arguments(tokens, parsed.end_index, len(tokens))
        if args is None:
            return None  # bang access or other unmodeled / malformed shape - leave raw
        return CallNode(span=span, has_call_keyword=False, callee=parsed.expr, args=args)

    def _parse_invalid_nested_module_block_statement(
        self, kind: str, stop_closers: Sequence[str]
    ) -> StatementNode:
        head = self._cursor.next()
        assert head is not None
        expected = "endtype" if kind == "type" else "endenum"
        end = head.end
        while not self._cursor.at_end():
            stmt = self._cursor.peek()
            assert stmt is not None
            closer = self._closer_kind(stmt)
            if closer == expected:
                nxt = self._cursor.next()
                assert nxt is not None
                end = nxt.end
                break
            if closer and closer in stop_closers:
                break
            if self._is_module_level_starter(stmt) and not (
                kind == "type" and self._is_type_field_statement(stmt)
            ):
                break
            nxt = self._cursor.next()
            assert nxt is not None
            end = nxt.end
        return StatementNode(span=Span(head.start, end), raw=self._source[head.start : end])

    def _parse_block(self, opener: str) -> BodyNode:
        head = self._cursor.next()
        assert head is not None
        expected = self._block_closer(opener)
        self._open_stack.append(expected)
        body: list[BodyNode] = []
        # If blocks accumulate structured arms; the flat body is kept for generic
        # body walkers and block-balance diagnostics.
        branches: list[_IfBranchBuilder] | None = (
            [self._start_if_branch(IfBranchKind.IF, head)] if opener == "if" else None
        )
        closed = False
        end_stmt: LogicalStatement | None = None
        while not self._cursor.at_end():
            stmt = self._cursor.peek()
            assert stmt is not None
            ck = self._closer_kind(stmt)
            if ck == expected:
                end_stmt = self._cursor.next()
                closed = True
                break
            nested = self._nested_type_or_enum_block_kind(stmt)
            if nested is not None:
                node = self._parse_invalid_nested_module_block_statement(nested, [expected])
                body.append(node)
                if branches is not None:
                    branches[len(branches) - 1].body.append(node)
                continue
            if self._is_module_level_starter(stmt):
                break
            if branches is not None:
                marker = self._if_branch_marker(stmt)
                if marker is not None:
                    self._cursor.next()
                    body.append(self._make_statement(stmt))  # keep header line in flat body
                    branches.append(self._start_if_branch(marker, stmt))
                    continue
            item = self._parse_body_item(stmt)
            if item is None:
                break
            body.append(item)
            if branches is not None:
                branches[len(branches) - 1].body.append(item)
        self._open_stack.pop()
        if not closed:
            self._diag(head, f"Block is missing {_CLOSER_LABELS[expected]}.", ParseSeverity.ERROR, "MS-VBAL 5.4")
        span = Span(head.start, (end_stmt if end_stmt is not None else head).end)
        if opener == "if":
            return self._finish_if_block(
                branches if branches is not None else [self._start_if_branch(IfBranchKind.IF, head)],
                body,
                closed,
                span,
            )
        return self._make_block_node(opener, body, closed, span, head, end_stmt)

    def _make_block_node(
        self,
        opener: str,
        body: list[BodyNode],
        closed: bool,
        span: Span,
        head: LogicalStatement,
        end_stmt: LogicalStatement | None,
    ) -> BodyNode:
        if opener == "for" or opener == "foreach":
            control = self._for_control_variable(opener, head)
            source = self._for_each_source_expression(head) if opener == "foreach" else None
            nxt = self._next_control_variable(end_stmt)
            for_node = ForBlockNode(span=span, each=(opener == "foreach"), closed=closed, body=body)
            if control is not None:
                for_node.control_variable, for_node.control_variable_span = control
            if source is not None:
                for_node.source_expression, for_node.source_expression_span = source
            if nxt is not None:
                for_node.next_variable, for_node.next_variable_span = nxt
            return for_node
        if opener == "do":
            return DoBlockNode(span=span, closed=closed, body=body)
        if opener == "while":
            return WhileBlockNode(span=span, closed=closed, body=body)
        if opener == "with":
            return WithBlockNode(span=span, closed=closed, body=body)
        if opener == "select":
            return SelectBlockNode(span=span, closed=closed, body=body)
        raise AssertionError(f"unreachable block opener: {opener}")

    def _if_branch_marker(self, stmt: LogicalStatement) -> IfBranchKind | None:
        tokens = _code_tokens_after_line_number(stmt)
        head = token_word(_at(tokens, 0))
        if head == "elseif":
            return IfBranchKind.ELSE_IF
        # `Else If` is the same arm as `ElseIf`; a bare `Else` is the else arm.
        if head == "else":
            return IfBranchKind.ELSE_IF if token_word(_at(tokens, 1)) == "if" else IfBranchKind.ELSE
        return None

    def _start_if_branch(self, branch_kind: IfBranchKind, stmt: LogicalStatement) -> _IfBranchBuilder:
        header_span = Span(stmt.start, stmt.end)
        if branch_kind is IfBranchKind.ELSE:
            return _IfBranchBuilder(branch_kind=branch_kind, condition=None, header_span=header_span)
        condition, condition_raw, condition_span = self._parse_then_condition(stmt, branch_kind)
        return _IfBranchBuilder(
            branch_kind=branch_kind,
            condition=condition,
            header_span=header_span,
            condition_raw=condition_raw,
            condition_span=condition_span,
        )

    def _parse_then_condition(
        self, stmt: LogicalStatement, branch_kind: IfBranchKind
    ) -> tuple[ExprNode | None, str | None, Span | None]:
        tokens = _code_tokens_after_line_number(stmt)
        # `If`/`ElseIf` is one token; `Else If` is two. Condition starts after them.
        cond_start = 2 if (branch_kind is IfBranchKind.ELSE_IF and token_word(_at(tokens, 0)) == "else") else 1
        then_index = -1
        for i in range(len(tokens) - 1, cond_start - 1, -1):
            t = tokens[i]
            if t.kind is TokenKind.KEYWORD and token_word(t) == "then":
                then_index = i
                break
        cond_end = then_index if then_index >= 0 else len(tokens)
        if cond_end <= cond_start:
            return (None, None, None)
        result = parse_expression(tokens, cond_start, cond_end)
        condition_span = Span(tokens[cond_start].start, tokens[cond_end - 1].end)
        condition = result.expr if _fully_consumed(result, cond_end) else None
        condition_raw = self._source[condition_span.start : condition_span.end]
        return (condition, condition_raw, condition_span)

    def _finish_if_block(
        self, builders: list[_IfBranchBuilder], body: list[BodyNode], closed: bool, span: Span
    ) -> IfBlockNode:
        branches: list[IfBranchNode] = []
        for i, builder in enumerate(builders):
            end = builders[i + 1].header_span.start if i + 1 < len(builders) else span.end
            branches.append(
                IfBranchNode(
                    branch_kind=builder.branch_kind,
                    condition=builder.condition,
                    body=builder.body,
                    header_span=builder.header_span,
                    span=Span(builder.header_span.start, end),
                    condition_raw=builder.condition_raw,
                    condition_span=builder.condition_span,
                )
            )
        return IfBlockNode(span=span, closed=closed, branches=branches, body=body)

    def _for_control_variable(
        self, opener: str, stmt: LogicalStatement
    ) -> tuple[str, Span] | None:
        tokens = _code_tokens_after_line_number(stmt)
        index = 2 if opener == "foreach" else 1
        name_token = _at(tokens, index)
        name = self._simple_name_from_token(name_token)
        if not name:
            return None
        if opener == "foreach":
            if token_word(_at(tokens, index + 1)) != "in":
                return None
        else:
            nt = _at(tokens, index + 1)
            if nt is None or nt.raw_text != "=":
                return None
        assert name_token is not None
        return (name, Span(name_token.start, name_token.end))

    def _for_each_source_expression(self, stmt: LogicalStatement) -> tuple[str, Span] | None:
        tokens = _code_tokens_after_line_number(stmt)
        in_index = 3 if token_word(_at(tokens, 3)) == "in" else -1
        if in_index < 0 or in_index + 1 >= len(tokens):
            return None
        first = tokens[in_index + 1]
        last = tokens[len(tokens) - 1]
        return (self._source[first.start : last.end], Span(first.start, last.end))

    def _next_control_variable(self, stmt: LogicalStatement | None) -> tuple[str, Span] | None:
        if stmt is None:
            return None
        tokens = _code_tokens_after_line_number(stmt)
        if token_word(_at(tokens, 0)) != "next" or len(tokens) != 2:
            return None
        name_token = tokens[1]
        name = self._simple_name_from_token(name_token)
        return (name, Span(name_token.start, name_token.end)) if name else None

    def _simple_name_from_token(self, token: VbaToken | None) -> str | None:
        if token is None:
            return None
        if token.kind is TokenKind.IDENTIFIER or token.kind is TokenKind.KEYWORD:
            return token.raw_text
        if token.kind is TokenKind.BRACKETED_IDENTIFIER:
            return self._strip_brackets(token.raw_text)
        return None

    @staticmethod
    def _block_closer(opener: str) -> str:
        return _BLOCK_CLOSERS[opener]

    def _opener_kind(self, stmt: LogicalStatement) -> str | None:
        tokens = _code_tokens_after_line_number(stmt)
        w0 = token_word(_at(tokens, 0))
        if w0 == "if":
            # Multi-line If only when Then is the final code token.
            last = _at(tokens, len(tokens) - 1)
            return "if" if token_word(last) == "then" else None
        if w0 == "for":
            return "foreach" if token_word(_at(tokens, 1)) == "each" else "for"
        if w0 == "do":
            return "do"
        if w0 == "while":
            return "while"
        if w0 == "with":
            return "with"
        if w0 == "select":
            return "select" if token_word(_at(tokens, 1)) == "case" else None
        return None

    def _closer_kind(self, stmt: LogicalStatement) -> str | None:
        tokens = _code_tokens_after_line_number(stmt)
        w0 = token_word(_at(tokens, 0))
        if w0 == "next":
            return "next"
        if w0 == "loop":
            return "loop"
        if w0 == "wend":
            return "wend"
        if w0 == "end":
            # "End" alone (MS-VBAL 5.4.7) is a statement, not a block closer.
            return _END_CLOSERS.get(token_word(_at(tokens, 1)))
        return None

    def _nested_type_or_enum_block_kind(self, stmt: LogicalStatement) -> str | None:
        tokens = _code_tokens_after_line_number(stmt)
        mod_index = self._leading_modifier_count(tokens)
        head = token_word(_at(tokens, mod_index))
        return head if (head == "type" or head == "enum") else None

    def _is_module_level_starter(self, stmt: LogicalStatement) -> bool:
        tokens = code_tokens(stmt)
        mod_index = self._leading_modifier_count(tokens)
        head = token_word(_at(tokens, mod_index))
        if head in ("sub", "function", "property", "type", "enum", "declare"):
            return True
        return token_word(_at(tokens, 0)) == "attribute"

    def _is_exported_procedure_attribute(
        self,
        stmt: LogicalStatement,
        tokens: Sequence[VbaToken],
        procedure_name: str,
        in_member_metadata_slot: bool,
    ) -> bool:
        if not self._starts_at_physical_line_start(stmt):
            return False
        if not in_member_metadata_slot:
            return False
        eq_index = next((i for i, t in enumerate(tokens) if t.raw_text == "="), -1)
        if eq_index <= 1:
            return False
        attr_name = "".join(t.raw_text for t in tokens[1:eq_index])
        dot = attr_name.find(".")
        if dot <= 0:
            return False
        target = self._strip_brackets(attr_name[:dot])
        member_attribute_name = attr_name[dot + 1 :]
        return (
            target.lower() == procedure_name.lower()
            and _VB_MEMBER_ATTR_RE.match(member_attribute_name) is not None
        )

    def _starts_at_physical_line_start(self, stmt: LogicalStatement) -> bool:
        previous_newline = max(
            self._source.rfind("\n", 0, stmt.start), self._source.rfind("\r", 0, stmt.start)
        )
        return stmt.start == previous_newline + 1

    # -- Token helpers -----------------------------------------------------

    @staticmethod
    def _leading_modifier_count(tokens: Sequence[VbaToken]) -> int:
        i = 0
        while True:
            t = _at(tokens, i)
            if t is None or token_word(t) not in _LEADING_MODIFIERS:
                break
            i += 1
        return i

    @staticmethod
    def _find_word_from(tokens: Sequence[VbaToken], word: str, from_index: int) -> int:
        for i in range(len(tokens)):
            if i >= from_index and token_word(tokens[i]) == word:
                return i
        return -1

    @staticmethod
    def _find_raw_after(tokens: Sequence[VbaToken], raw: str, after_index: int) -> int:
        for i in range(len(tokens)):
            if i > after_index and tokens[i].raw_text == raw:
                return i
        return -1

    @staticmethod
    def _split_top_level_commas(
        tokens: Sequence[VbaToken], from_: int, to: int
    ) -> list[list[VbaToken]]:
        return split_top_level_token_groups(tokens, from_, ",", to)

    def _directive_condition(
        self, tokens: Sequence[VbaToken], from_: int
    ) -> tuple[str | None, Span | None]:
        to = len(tokens) - 1 if token_word(_at(tokens, len(tokens) - 1)) == "then" else len(tokens)
        return self._token_range_raw(tokens, from_, to)

    def _token_range_raw(
        self, tokens: Sequence[VbaToken], from_: int, to: int
    ) -> tuple[str | None, Span | None]:
        if from_ < 0 or from_ >= to or from_ >= len(tokens):
            return (None, None)
        start_token = tokens[from_]
        end_token = tokens[min(to, len(tokens)) - 1]
        return (self._source[start_token.start : end_token.end], Span(start_token.start, end_token.end))

    @staticmethod
    def _match_paren(tokens: Sequence[VbaToken], lparen: int) -> int:
        close = match_paren_from(tokens, lparen)
        return close if close >= 0 else len(tokens) - 1

    def _skip_parens(self, tokens: Sequence[VbaToken], i: int) -> int:
        return self._match_paren(tokens, i) + 1

    def _capture_type(self, tokens: Sequence[VbaToken], i: int, stop_raw: str | None = None) -> str | None:
        if _at(tokens, i) is None:
            return None
        depth = 0
        last = i
        for j in range(i, len(tokens)):
            t = tokens[j]
            if t.raw_text == "(":
                depth += 1
            elif t.raw_text == ")":
                if depth == 0:
                    break
                depth -= 1
            if depth == 0 and (t.raw_text == "," or (stop_raw is not None and t.raw_text == stop_raw)):
                break
            last = j
        return self._source[tokens[i].start : tokens[last].end]

    def _capture_declaration_type(
        self, tokens: Sequence[VbaToken], i: int, stop_raw: str | None = None
    ) -> tuple[str | None, str | None]:
        fixed = parse_fixed_length_string_type(tokens, i)
        if fixed is not None:
            as_type = self._source[tokens[i].start : tokens[i].end]
            fixed_length = self._source[
                tokens[fixed.length_index].start : tokens[fixed.length_index].end
            ]
            return (as_type, fixed_length)
        return (self._capture_type(tokens, i, stop_raw), None)

    @staticmethod
    def _canonical(token: VbaToken | None) -> str:
        if token is None:
            return ""
        return token.canonical_text if token.canonical_text is not None else token.raw_text

    def _parse_declared_name(
        self, tokens: Sequence[VbaToken], index: int, allow_type_suffix: bool
    ) -> _DeclaredNameInfo:
        name_token = _at(tokens, index)
        name = self._strip_brackets(name_token.raw_text) if name_token is not None else ""
        name_span = self._token_span(name_token)
        suffix_token = _at(tokens, index + 1) if allow_type_suffix else None
        if (
            name_token is not None
            and suffix_token is not None
            and name_token.end == suffix_token.start
            and is_type_declaration_suffix(suffix_token.raw_text)
        ):
            return _DeclaredNameInfo(
                name=name,
                next_index=index + 2,
                name_span=name_span,
                type_suffix=suffix_token.raw_text,
                type_suffix_span=Span(suffix_token.start, suffix_token.end),
            )
        return _DeclaredNameInfo(name=name, next_index=index + 1, name_span=name_span)

    @staticmethod
    def _token_span(token: VbaToken | None) -> Span | None:
        return Span(token.start, token.end) if token is not None else None

    @staticmethod
    def _strip_brackets(raw: str) -> str:
        if len(raw) >= 2 and raw.startswith("[") and raw.endswith("]"):
            return raw[1:-1]
        return raw

    @staticmethod
    def _string_literal_text(token: VbaToken | None) -> str | None:
        if token is None or token.kind is not TokenKind.STRING_LITERAL:
            return None
        raw = token.raw_text
        if len(raw) >= 2 and raw.startswith('"') and raw.endswith('"'):
            return raw[1:-1].replace('""', '"')
        return raw

    def _make_statement(self, stmt: LogicalStatement) -> StatementNode:
        return StatementNode(span=Span(stmt.start, stmt.end), raw=self._source[stmt.start : stmt.end])

    @staticmethod
    def _detect_module_kind(members: Sequence[ModuleMember]) -> ModuleKind:
        for m in members:
            if isinstance(m, AttributeNode) and _VB_CLASS_ATTR_RE.match(m.name) is not None:
                return ModuleKind.CLASS
        return ModuleKind.UNKNOWN

    def _diag(
        self, at: LogicalStatement, message: str, severity: ParseSeverity, spec_ref: str | None = None
    ) -> None:
        self._diagnostics.append(
            ParseDiagnostic(span=Span(at.start, at.end), message=message, severity=severity, spec_ref=spec_ref)
        )


def _code_tokens_after_line_number(statement: LogicalStatement) -> list[VbaToken]:
    return tokens_without_leading_line_number(code_tokens(statement))


def _top_level_equals_index(tokens: Sequence[VbaToken], from_: int) -> int:
    """Index of the first depth-0 '=' assignment operator at or after from_, or -1.

    `<=`/`>=`/`<>`/`:=` are distinct tokens, and an '=' inside parentheses is a
    comparison, so neither is matched (MS-VBAL 5.4.3).
    """
    depth = 0
    for i in range(from_, len(tokens)):
        raw = tokens[i].raw_text
        if raw == "(" or raw == "[":
            depth += 1
        elif raw == ")" or raw == "]":
            depth -= 1
        elif depth == 0 and tokens[i].kind is TokenKind.OPERATOR and raw == "=":
            return i
    return -1


def _is_mid_statement_target(tokens: Sequence[VbaToken], start: int) -> bool:
    """True when the LHS at start is Mid(/Mid$(/MidB(/MidB$( - the dedicated Mid
    replacement statement (MS-VBAL 5.4.3.x), not a generic assignment."""
    word = token_word(_at(tokens, start))
    if word != "mid" and word != "midb":
        return False
    after = _at(tokens, start + 1)
    if after is not None and after.raw_text == "(":
        return True
    two = _at(tokens, start + 2)
    return after is not None and after.raw_text == "$" and two is not None and two.raw_text == "("


def _fully_consumed(result: ExprParseResult, to: int) -> bool:
    """True when an expression parse cleanly consumed exactly tokens[:to]: a
    non-null expression, stopped at the boundary, and zero diagnostics."""
    return result.expr is not None and result.end_index == to and len(result.diagnostics) == 0
