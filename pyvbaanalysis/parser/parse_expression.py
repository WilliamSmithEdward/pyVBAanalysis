"""Error-tolerant VBA value-expression parser (MS-VBAL 5.6).

Ported from xlide_vscode/src/analyzer/parser/parseExpression.ts. Turns a slice of
significant tokens into the ExprNode hierarchy in nodes.py.

Design notes:
- Never throws. On an unexpected token the parser stops at the last good position,
  records a diagnostic, and returns a best-effort node.
- Every node carries an absolute source span built from token offsets.
- Scope: literals, identifiers, parenthesised expressions, member-access chains
  (including the leading-dot With form), index/call expressions with positional,
  named (name:=expr), and omitted (f(1, , 3)) arguments, unary -/+/Not, the full
  binary precedence ladder, New / AddressOf / TypeOf...Is, and bang access (obj!name).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..lexer.token_helpers import is_ident_like, token_name, token_word
from ..lexer.token_kinds import TokenKind, VbaToken
from .nodes import (
    AddressOfExpr,
    Argument,
    BinaryExpr,
    ExprNode,
    IdentifierExpr,
    IndexExpr,
    LiteralExpr,
    LiteralKind,
    MemberAccessExpr,
    MemberAccessKind,
    NewExpr,
    ParenExpr,
    ParseDiagnostic,
    ParseSeverity,
    Span,
    TypeOfIsExpr,
    UnaryExpr,
)


@dataclass(frozen=True, slots=True)
class ExprParseResult:
    """Result of parsing an expression from a token slice."""

    # The parsed expression, or None when the slice held no expression token.
    expr: ExprNode | None
    # Diagnostics raised while parsing (unexpected/trailing tokens, etc.).
    diagnostics: list[ParseDiagnostic]
    # Index into the input token array just past the last consumed token.
    end_index: int


# Binary-operator binding power (MS-VBAL 5.6.6). Higher binds tighter. All binary
# operators here are left-associative; ^ and the unary operators are handled
# structurally in _parse_unary/_parse_power because ^ binds tighter than unary
# minus while still allowing a signed exponent.
BINARY_PRECEDENCE: dict[str, int] = {
    "Imp": 1,
    "Eqv": 2,
    "Xor": 3,
    "Or": 4,
    "And": 5,
    "=": 6,
    "<>": 6,
    "<": 6,
    ">": 6,
    "<=": 6,
    ">=": 6,
    "Like": 6,
    "Is": 6,
    "&": 7,
    "+": 8,
    "-": 8,
    "Mod": 9,
    "\\": 10,
    "*": 11,
    "/": 11,
}

# Word-operators that arrive as keyword tokens, mapped to canonical text.
WORD_BINARY_OPS: dict[str, str] = {
    "mod": "Mod",
    "and": "And",
    "or": "Or",
    "xor": "Xor",
    "eqv": "Eqv",
    "imp": "Imp",
    "like": "Like",
    "is": "Is",
}

LITERAL_KEYWORDS: dict[str, LiteralKind] = {
    "true": LiteralKind.BOOLEAN,
    "false": LiteralKind.BOOLEAN,
    "nothing": LiteralKind.NOTHING,
    "null": LiteralKind.NULL,
    "empty": LiteralKind.EMPTY,
}

# Type-declaration characters that are never binary operators (safe as suffix).
UNAMBIGUOUS_SUFFIXES: frozenset[str] = frozenset({"$", "%", "@"})

# Binding power of the Not prefix (MS-VBAL 5.6.6): below the comparison operators
# but above And, so `Not a = b` is `Not (a = b)` while `Not a And b` is
# `(Not a) And b`.
NOT_PRECEDENCE = 6


def parse_expression(
    tokens: Sequence[VbaToken], from_: int = 0, to: int | None = None
) -> ExprParseResult:
    """Parse a value expression from tokens[from_:to).

    The token slice must already exclude comments and statement separators.
    """
    end = len(tokens) if to is None else to
    parser = _ExpressionParser(tokens, from_, end)
    expr = parser.parse()
    return ExprParseResult(expr=expr, diagnostics=parser.diagnostics, end_index=parser.index)


def parse_parenless_arguments(
    tokens: Sequence[VbaToken], from_: int, to: int
) -> list[Argument] | None:
    """Parse a parenless call-statement argument list from tokens[from_:to).

    Returns None when any present argument is malformed (the caller falls back to a
    raw statement). Supports positional, named (name:=value), and omitted arguments.
    """
    return _ExpressionParser(tokens, from_, to).parse_parenless_argument_list()


class _ExpressionParser:
    __slots__ = ("_tokens", "_to", "index", "diagnostics")

    def __init__(self, tokens: Sequence[VbaToken], from_: int, to: int) -> None:
        self._tokens = tokens
        self._to = to
        self.index = from_
        self.diagnostics: list[ParseDiagnostic] = []

    def parse(self) -> ExprNode | None:
        if self._at_end():
            return None
        return self._parse_binary(0)

    # --- token cursor -------------------------------------------------------

    def _at_end(self) -> bool:
        return self.index >= self._to

    def _peek(self) -> VbaToken | None:
        return self._tokens[self.index] if self.index < self._to else None

    def _next(self) -> VbaToken | None:
        if self.index < self._to:
            tok = self._tokens[self.index]
            self.index += 1
            return tok
        return None

    def _peek_at(self, offset: int) -> VbaToken | None:
        i = self.index + offset
        return self._tokens[i] if i < self._to else None

    @staticmethod
    def _raw(token: VbaToken | None) -> str:
        return token.raw_text if token is not None else ""

    @staticmethod
    def _name_or_raw(token: VbaToken) -> str:
        name = token_name(token)
        return name if name is not None else token.raw_text

    @staticmethod
    def _span_of(expr: ExprNode) -> Span:
        return expr.span

    def _diag(self, token: VbaToken | None, message: str) -> None:
        span = Span(token.start, token.end) if token is not None else Span(0, 0)
        self.diagnostics.append(
            ParseDiagnostic(span=span, message=message, severity=ParseSeverity.ERROR, spec_ref="MS-VBAL 5.6")
        )

    # --- precedence-climbing binary layer -----------------------------------

    def _parse_binary(self, min_prec: int) -> ExprNode | None:
        left: ExprNode | None
        # `Not` is a low-precedence prefix: recognise it only where a Not-expression
        # is allowed (at or below its binding power).
        lead = self._peek()
        if (
            lead is not None
            and lead.kind is TokenKind.KEYWORD
            and token_word(lead) == "not"
            and min_prec <= NOT_PRECEDENCE
        ):
            self._next()
            operand = self._parse_binary(NOT_PRECEDENCE)
            if operand is None:
                self._diag(lead, "Expected an expression after 'Not'.")
                return None
            left = UnaryExpr(
                span=Span(lead.start, self._span_of(operand).end), operator="Not", operand=operand
            )
        else:
            left = self._parse_unary()
        if left is None:
            return None
        while True:
            op_token = self._peek()
            op = self._binary_operator(op_token)
            if op is None:
                break
            prec = BINARY_PRECEDENCE.get(op)
            if prec is None or prec < min_prec:
                break
            self._next()  # consume operator
            # Left-associative: the right side binds operators strictly tighter.
            right = self._parse_binary(prec + 1)
            if right is None:
                self._diag(op_token, f"Expected an expression after '{op}'.")
                return left
            left = BinaryExpr(
                span=Span(self._span_of(left).start, self._span_of(right).end),
                operator=op,
                left=left,
                right=right,
            )
        return left

    def _binary_operator(self, token: VbaToken | None) -> str | None:
        """Canonical binary operator for a token, or None when it is not one."""
        if token is None:
            return None
        if token.kind is TokenKind.KEYWORD:
            return WORD_BINARY_OPS.get(token_word(token))
        if token.kind is TokenKind.OPERATOR:
            raw = token.raw_text
            if raw in BINARY_PRECEDENCE:
                return raw
        return None

    # --- unary / exponent ---------------------------------------------------

    def _parse_unary(self) -> ExprNode | None:
        token = self._peek()
        if token is None:
            self._diag(None, "Expected an expression.")
            return None
        unary_op = self._prefix_operator(token)
        if unary_op is not None:
            self._next()
            operand = self._parse_unary()
            if operand is None:
                self._diag(token, f"Expected an expression after '{self._raw(token)}'.")
                return None
            return UnaryExpr(
                span=Span(token.start, self._span_of(operand).end), operator=unary_op, operand=operand
            )
        return self._parse_power()

    @staticmethod
    def _prefix_operator(token: VbaToken) -> str | None:
        """Prefix - or + (MS-VBAL 5.6.6). Not is handled in _parse_binary."""
        if token.kind is TokenKind.OPERATOR and (token.raw_text == "-" or token.raw_text == "+"):
            return token.raw_text
        return None

    def _parse_power(self) -> ExprNode | None:
        base = self._parse_postfix()
        if base is None:
            return None
        while True:
            p = self._peek()
            if p is None or p.kind is not TokenKind.OPERATOR or p.raw_text != "^":
                break
            op_token = self._next()
            exponent = self._parse_signed_primary()
            if exponent is None:
                self._diag(op_token, "Expected an expression after '^'.")
                return base
            base = BinaryExpr(
                span=Span(self._span_of(base).start, self._span_of(exponent).end),
                operator="^",
                left=base,
                right=exponent,
            )
        return base

    def _parse_signed_primary(self) -> ExprNode | None:
        """A postfix primary optionally preceded by sign(s) - the operand of ^."""
        token = self._peek()
        if token is not None and token.kind is TokenKind.OPERATOR and (token.raw_text == "-" or token.raw_text == "+"):
            self._next()
            operand = self._parse_signed_primary()
            if operand is None:
                return None
            return UnaryExpr(
                span=Span(token.start, self._span_of(operand).end),
                operator=token.raw_text,
                operand=operand,
            )
        return self._parse_postfix()

    # --- postfix: member access and index/call ------------------------------

    def _parse_postfix(self) -> ExprNode | None:
        expr = self._parse_primary()
        if expr is None:
            return None
        while True:
            token = self._peek()
            if token is None:
                break
            if token.kind is TokenKind.PUNCTUATION and token.raw_text == ".":
                self._next()
                member_token = self._peek()
                if member_token is None or not self._is_member_name(member_token):
                    self._diag(member_token if member_token is not None else token, "Expected a member name after '.'.")
                    break
                self._next()
                member = self._name_or_raw(member_token)
                expr = MemberAccessExpr(
                    span=Span(self._span_of(expr).start, member_token.end),
                    object_=expr,
                    member=member,
                    member_span=Span(member_token.start, member_token.end),
                )
                continue
            if token.kind is TokenKind.PUNCTUATION and token.raw_text == "(":
                indexed = self._parse_index(expr)
                if indexed is None:
                    break
                expr = indexed
                continue
            # Bang member access: receiver!name / receiver![Bracketed Name].
            bang = self._bang_member_access(expr, token)
            if bang is not None:
                expr = bang
                continue
            break
        return expr

    def _bang_member_access(self, receiver: ExprNode, bang_token: VbaToken) -> MemberAccessExpr | None:
        """Build a receiver!name bang node, or None when ! is not a bang."""
        if bang_token.kind is not TokenKind.OPERATOR or bang_token.raw_text != "!":
            return None
        name_tok = self._peek_at(1)
        if (
            name_tok is None
            or not (name_tok.kind is TokenKind.IDENTIFIER or name_tok.kind is TokenKind.BRACKETED_IDENTIFIER)
            or bang_token.start != self._span_of(receiver).end  # ! glued to receiver
            or name_tok.start != bang_token.end  # name glued to !
        ):
            return None
        self._next()  # consume !
        self._next()  # consume name
        member = self._name_or_raw(name_tok)
        return MemberAccessExpr(
            span=Span(self._span_of(receiver).start, name_tok.end),
            object_=receiver,
            member=member,
            member_span=Span(name_tok.start, name_tok.end),
            access_kind=MemberAccessKind.BANG,
        )

    def _parse_index(self, callee: ExprNode) -> IndexExpr | None:
        """callee(args) - positional, named (name:=expr), and omitted arguments."""
        open_tok = self._next()  # consume '('
        assert open_tok is not None
        args: list[Argument] = []
        head = self._peek()
        if head is not None and head.kind is TokenKind.PUNCTUATION and head.raw_text == ")":
            close = self._next()
            assert close is not None
            return IndexExpr(span=Span(self._span_of(callee).start, close.end), callee=callee, args=args)
        while True:
            arg = self._parse_argument(")")
            if arg is None:
                cur = self._peek()
                self._diag(cur if cur is not None else open_tok, "Expected an argument expression.")
                return None
            args.append(arg)
            sep = self._peek()
            if sep is not None and sep.kind is TokenKind.PUNCTUATION and sep.raw_text == ",":
                self._next()
                continue
            if sep is not None and sep.kind is TokenKind.PUNCTUATION and sep.raw_text == ")":
                close = self._next()
                assert close is not None
                return IndexExpr(span=Span(self._span_of(callee).start, close.end), callee=callee, args=args)
            self._diag(sep if sep is not None else open_tok, "Expected ',' or ')' in argument list.")
            return None

    def parse_parenless_argument_list(self) -> list[Argument] | None:
        """Parse a parenless argument list from the remaining tokens.

        Returns None when any present argument is malformed so the caller can fall
        back to a raw statement; an empty trailing slot after a comma (Foo a,) is
        treated as malformed rather than an omission.
        """
        args: list[Argument] = []
        while True:
            arg = self._parse_argument(None)
            if arg is None:
                return None
            args.append(arg)
            sep = self._peek()
            if sep is None:
                return args  # reached the end of the slice
            if sep.kind is TokenKind.PUNCTUATION and sep.raw_text == ",":
                self._next()
                continue
            return None  # unexpected trailing token - malformed

    def _parse_argument(self, terminator: str | None) -> Argument | None:
        """Parse a single argument: an optional name:= prefix then a value, or an
        omitted slot. Returns None only when a value is expected but fails."""
        head = self._peek()
        # Omitted slot: empty position before a separator or terminator.
        if self._at_argument_boundary(head, terminator):
            assert head is not None
            pos = head.start
            return Argument(value=None, span=Span(pos, pos))
        if head is None:
            return None
        # Named argument: name := value.
        name: str | None = None
        name_span: Span | None = None
        colon_eq = self._peek_at(1)
        if (
            self._is_member_name(head)
            and colon_eq is not None
            and colon_eq.kind is TokenKind.OPERATOR
            and colon_eq.raw_text == ":="
        ):
            name = self._name_or_raw(head)
            name_span = Span(head.start, head.end)
            self._next()  # consume name
            self._next()  # consume :=
        value = self._parse_binary(0)
        if value is None:
            return None
        start = name_span.start if name_span is not None else self._span_of(value).start
        return Argument(value=value, span=Span(start, self._span_of(value).end), name=name, name_span=name_span)

    @staticmethod
    def _at_argument_boundary(token: VbaToken | None, terminator: str | None) -> bool:
        """True when the current position is an empty argument slot (, or terminator)."""
        if token is None or token.kind is not TokenKind.PUNCTUATION:
            return False
        return token.raw_text == "," or (terminator is not None and token.raw_text == terminator)

    @staticmethod
    def _is_member_name(token: VbaToken) -> bool:
        return is_ident_like(token) or token.kind is TokenKind.BRACKETED_IDENTIFIER

    # --- primary ------------------------------------------------------------

    def _parse_primary(self) -> ExprNode | None:
        token = self._peek()
        if token is None:
            self._diag(None, "Expected an expression.")
            return None

        # Parenthesised expression.
        if token.kind is TokenKind.PUNCTUATION and token.raw_text == "(":
            self._next()
            inner = self._parse_binary(0)
            if inner is None:
                return None
            close = self._peek()
            if close is not None and close.kind is TokenKind.PUNCTUATION and close.raw_text == ")":
                self._next()
                return ParenExpr(span=Span(token.start, close.end), inner=inner)
            self._diag(close if close is not None else token, "Expected ')'.")
            return inner

        # Leading-dot member access inside a With block (.Member).
        if token.kind is TokenKind.PUNCTUATION and token.raw_text == ".":
            self._next()
            member_token = self._peek()
            if member_token is None or not self._is_member_name(member_token):
                self._diag(member_token if member_token is not None else token, "Expected a member name after '.'.")
                return None
            self._next()
            member = self._name_or_raw(member_token)
            return MemberAccessExpr(
                span=Span(token.start, member_token.end),
                object_=None,
                member=member,
                member_span=Span(member_token.start, member_token.end),
            )

        # Literals.
        literal = self._literal_for(token)
        if literal is not None:
            self._next()
            return literal

        # Keyword-led primaries: New / AddressOf / TypeOf.
        if token.kind is TokenKind.KEYWORD:
            word = token_word(token)
            if word == "new":
                return self._parse_new(token)
            if word == "addressof":
                return self._parse_address_of(token)
            if word == "typeof":
                return self._parse_type_of(token)

        # Identifier reference.
        if token.kind is TokenKind.IDENTIFIER or token.kind is TokenKind.BRACKETED_IDENTIFIER:
            self._next()
            return self._identifier_expr(token)

        self._diag(token, f"Unexpected token '{token.raw_text}' in expression.")
        return None

    def _identifier_expr(self, token: VbaToken) -> IdentifierExpr:
        name = self._name_or_raw(token)
        node = IdentifierExpr(span=Span(token.start, token.end), name=name)
        # Capture an immediately-adjacent unambiguous type-declaration character
        # ($/%/@). &/!/#/^ are left to the operator layer to avoid misreading a
        # concatenation, bang, date, or exponent.
        suffix = self._peek()
        if (
            suffix is not None
            and suffix.kind is TokenKind.UNKNOWN
            and suffix.raw_text in UNAMBIGUOUS_SUFFIXES
            and suffix.start == token.end
        ):
            self._next()
            node.type_suffix = suffix.raw_text
            node.span = Span(token.start, suffix.end)
        return node

    @staticmethod
    def _literal_for(token: VbaToken) -> LiteralExpr | None:
        literal_kind: LiteralKind | None
        kind = token.kind
        if kind is TokenKind.INTEGER_LITERAL:
            literal_kind = LiteralKind.INTEGER
        elif kind is TokenKind.FLOAT_LITERAL:
            literal_kind = LiteralKind.FLOAT
        elif kind is TokenKind.STRING_LITERAL:
            literal_kind = LiteralKind.STRING
        elif kind is TokenKind.DATE_LITERAL:
            literal_kind = LiteralKind.DATE
        elif kind is TokenKind.KEYWORD:
            literal_kind = LITERAL_KEYWORDS.get(token_word(token))
        else:
            literal_kind = None
        if literal_kind is None:
            return None
        return LiteralExpr(span=Span(token.start, token.end), literal_kind=literal_kind, raw=token.raw_text)

    def _parse_new(self, keyword: VbaToken) -> NewExpr | None:
        self._next()  # consume New
        type_token = self._peek()
        if type_token is None or not self._is_member_name(type_token):
            self._diag(type_token if type_token is not None else keyword, "Expected a type name after 'New'.")
            return None
        # A New type name may be a dotted library type (e.g. Scripting.Dictionary).
        end_token = self._next()
        assert end_token is not None
        type_name = self._name_or_raw(end_token)
        while True:
            p = self._peek()
            if p is None or p.kind is not TokenKind.PUNCTUATION or p.raw_text != ".":
                break
            self._next()
            part = self._peek()
            if part is None or not self._is_member_name(part):
                break
            self._next()
            type_name += f".{self._name_or_raw(part)}"
            end_token = part
        return NewExpr(
            span=Span(keyword.start, end_token.end),
            type_name=type_name,
            type_name_span=Span(keyword.start, end_token.end),
        )

    def _parse_address_of(self, keyword: VbaToken) -> AddressOfExpr | None:
        self._next()  # consume AddressOf
        target_token = self._peek()
        if target_token is None or not (target_token.kind is TokenKind.IDENTIFIER or is_ident_like(target_token)):
            self._diag(target_token if target_token is not None else keyword, "Expected a procedure name after 'AddressOf'.")
            return None
        self._next()
        target = self._identifier_expr(target_token)
        return AddressOfExpr(span=Span(keyword.start, target.span.end), target=target)

    def _parse_type_of(self, keyword: VbaToken) -> TypeOfIsExpr | None:
        self._next()  # consume TypeOf
        # Operand runs up to the Is keyword; parse it as a postfix expression.
        operand = self._parse_postfix()
        if operand is None:
            self._diag(keyword, "Expected an expression after 'TypeOf'.")
            return None
        is_token = self._peek()
        if is_token is None or is_token.kind is not TokenKind.KEYWORD or token_word(is_token) != "is":
            self._diag(is_token if is_token is not None else keyword, "Expected 'Is' in a 'TypeOf ... Is' expression.")
            return None
        self._next()  # consume Is
        type_token = self._peek()
        if type_token is None or not self._is_member_name(type_token):
            self._diag(type_token if type_token is not None else is_token, "Expected a type name after 'Is'.")
            return None
        self._next()
        type_name = self._name_or_raw(type_token)
        return TypeOfIsExpr(
            span=Span(keyword.start, type_token.end),
            operand=operand,
            type_name=type_name,
            type_name_span=Span(type_token.start, type_token.end),
        )
