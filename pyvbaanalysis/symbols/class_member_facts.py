"""What a class's own code shows a member always gives.

Ported from xlide_vscode/src/analyzer/symbols/classMemberFacts.ts (XLIDE issue
#414, each measured in Excel 16.0 through ``Dim c As New Class1``):

    Public M As Object, nothing in the class assigns it    Nothing
    Function M() As Object that only sets it to Nothing    Nothing
    Public M As Variant, nothing in the class assigns it   Empty
    Property Get M() As Variant: M = 1, and nothing else   a scalar
    Function M() As Variant: M = 1, and nothing else       a scalar
    Function or Get As Variant that never assigns M        Empty

Code outside the class can still assign a field through the instance; the rule
that reads these facts follows the instance's own uses.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from typing import Literal

from ..js_compat import JS_WHITESPACE, js_trim
from ..lexer.token_helpers import first_token_at_or_after
from ..lexer.token_kinds import TokenKind, VbaToken
from ..lexer.tokenize import tokenize_cached
from .symbol_model import VbaSymbol, VbaSymbolKind

_SCALAR_TYPES = frozenset(
    (
        "string", "boolean", "date", "byte", "integer", "long", "longlong", "longptr",
        "single", "double", "currency", "decimal",
    )
)

_WS = "[" + re.escape(JS_WHITESPACE) + "]"
_ARRAY_SUFFIX_RE = re.compile(_WS + r"*\(" + _WS + r"*\)" + _WS + r"*\Z")
_VBA_PREFIX_RE = re.compile(r"^vba\.", re.IGNORECASE)


def _normalize_type(type_: str | None) -> str | None:
    if type_ is None:
        return None
    text = _VBA_PREFIX_RE.sub("", _ARRAY_SUFFIX_RE.sub("", js_trim(type_), count=1), count=1).lower()
    return text or None


def _is_known_scalar_type(type_: str) -> bool:
    return type_ in _SCALAR_TYPES


ClassMemberValue = Literal["nothing", "empty", "scalar"]

# Words that make a body's flow more than one straight run.
_FLOW_WORDS = frozenset(
    ("if", "select", "for", "do", "while", "with", "goto", "gosub", "on", "exit", "resume", "end")
)

_LITERAL_KINDS = (TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL)


def _word(tok: VbaToken | None) -> str:
    return tok.raw_text.lower() if tok is not None else ""


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def class_member_values(source: str, children: Sequence[VbaSymbol]) -> dict[str, ClassMemberValue]:
    """The member values the class's code decides, by lowercased member name."""
    out: dict[str, ClassMemberValue] = {}
    toks = [tok for tok in tokenize_cached(source) if tok.kind is not TokenKind.COMMENT]
    # Only the earliest/latest mention is needed to decide whether a name
    # occurs outside its declaration. Build that index once, not per field.
    mentions_by_name: dict[str, list[int]] = {}
    for tok in toks:
        if tok.kind is not TokenKind.IDENTIFIER:
            continue
        lower = _word(tok)
        found = mentions_by_name.get(lower)
        if found is not None:
            found[1] = tok.start
        else:
            mentions_by_name[lower] = [tok.start, tok.start]
    for symbol in children:
        lower = symbol.name.lower()
        type_ = _normalize_type(symbol.as_type)
        if symbol.kind is VbaSymbolKind.MODULE_VARIABLE:
            if symbol.is_array or symbol.is_auto_instantiated:
                continue
            # Any mention past the declaration, Me.M and a ByRef pass included,
            # may assign it.
            first_last = mentions_by_name.get(lower)
            named = first_last is not None and (
                first_last[0] < symbol.full_span.start or first_last[1] >= symbol.full_span.end
            )
            if named:
                continue
            if type_ is None or type_ == "variant":
                out[lower] = "empty"
            elif not _is_known_scalar_type(type_):
                out[lower] = "nothing"
            continue
        if symbol.kind is not VbaSymbolKind.FUNCTION and symbol.kind is not VbaSymbolKind.PROPERTY_GET:
            continue
        if any(child.kind is VbaSymbolKind.PARAMETER for child in symbol.children or []):
            continue
        # The body's statements, the header line left out.
        body: list[VbaToken] = []
        for i in range(first_token_at_or_after(toks, symbol.name_span.end + 1), len(toks)):
            tok = toks[i]
            if tok.start > symbol.full_span.end:
                break
            if tok.end <= symbol.full_span.end:
                body.append(tok)
        header_end = next((i for i, tok in enumerate(body) if tok.kind is TokenKind.NEWLINE), -1)
        statements = _split_statements(body[header_end + 1 :])
        # The last statement is End Function or End Property.
        inner = [
            stmt
            for stmt in statements
            if not (_word(_at(stmt, 0)) == "end" and _word(_at(stmt, 1)) in ("function", "property"))
        ]
        mentions = [
            stmt
            for stmt in inner
            if any(
                tok.kind is TokenKind.IDENTIFIER
                and _word(tok) == lower
                and (i == 0 or stmt[i - 1].raw_text != ".")
                for i, tok in enumerate(stmt)
            )
        ]
        if type_ is not None and type_ != "variant" and not _is_known_scalar_type(type_):
            # An object result never set, or set only to Nothing, is Nothing.
            if all(
                len(stmt) == 4
                and _word(stmt[0]) == "set"
                and _word(stmt[1]) == lower
                and stmt[2].raw_text == "="
                and _word(stmt[3]) == "nothing"
                for stmt in mentions
            ):
                out[lower] = "nothing"
            continue
        # A Variant result nothing assigns is Empty (XLIDE issue #414, measured).
        if (type_ is None or type_ == "variant") and len(mentions) == 0:
            out[lower] = "empty"
            continue
        if (
            (type_ is None or type_ == "variant")
            and len(mentions) == 1
            and not any(_word(_at(stmt, 0)) in _FLOW_WORDS for stmt in inner)
        ):
            stmt = mentions[0]
            value = stmt[2:]
            literal = (
                value[-1]
                if len(value) == 1 or (len(value) == 2 and value[0].raw_text == "-")
                else None
            )
            second = _at(stmt, 1)
            if (
                _word(stmt[0]) == lower
                and second is not None
                and second.raw_text == "="
                and literal is not None
                and literal.kind in _LITERAL_KINDS
            ):
                out[lower] = "scalar"
    return out


def _split_statements(toks: Sequence[VbaToken]) -> list[list[VbaToken]]:
    """Tokens split at line ends and colons into statements, empty ones dropped."""
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    for tok in toks:
        if tok.kind is TokenKind.NEWLINE or tok.raw_text == ":":
            if current:
                out.append(current)
            current = []
            continue
        current.append(tok)
    if current:
        out.append(current)
    return out
