"""Small token utilities shared across analyzer surfaces.

Ported from xlide_vscode/src/analyzer/lexer/tokenHelpers.ts. Keeps statement-level
token handling (comment/newline filtering, identifier extraction, leading line
numbers, paren matching) from drifting between surfaces. VBA is case-insensitive
(MS-VBAL 3.3.5), so name matching folds case.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Sequence
from typing import Literal, Protocol

from ..identity_cache import IdentityLru
from .token_kinds import TokenKind, VbaToken
from .tokenize import tokenize, tokenize_cached

_DECIMAL_RE = re.compile(r"^\d+$")


def is_ident_like(token: VbaToken) -> bool:
    """True when the token reads as a bare identifier (identifier or keyword)."""
    return (
        token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
        and is_identifier(token.raw_text)
    )


def statement_tokens(source: str, start: int, end: int) -> list[VbaToken]:
    """Significant tokens of a source span, excluding comments and newlines."""
    return [
        t
        for t in tokenize(source[start:end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]


# Statement-token cache (audit #5): independent rules re-tokenize the same
# statement many times per pass. Tokens are cached per source string (LRU of 2)
# and per statement span, so one pass lexes each statement at most once - and
# usually not at all, because the span view is derived from the module's one
# shared tokenization. Callers must not mutate the returned lists.
_STATEMENT_TOKEN_CACHE_MAX = 2
_statement_token_cache: list[tuple[str, dict[tuple[int, int], list[VbaToken]]]] = []


def cached_statement_tokens(source: str, start: int, end: int) -> list[VbaToken]:
    """Memoized statement_tokens: identical output, one lex per statement per pass."""
    entry: dict[tuple[int, int], list[VbaToken]] | None = None
    for i, (cached_source, by_span) in enumerate(_statement_token_cache):
        if cached_source == source:
            entry = by_span
            if i > 0:
                _statement_token_cache.insert(0, _statement_token_cache.pop(i))
            break
    if entry is None:
        entry = {}
        _statement_token_cache.insert(0, (source, entry))
        if len(_statement_token_cache) > _STATEMENT_TOKEN_CACHE_MAX:
            _statement_token_cache.pop()
    key = (start, end)
    toks = entry.get(key)
    if toks is None:
        toks = _derive_statement_tokens(source, start, end)
        if toks is None:
            toks = statement_tokens(source, start, end)
        entry[key] = toks
    return toks


def cached_raw_statement_tokens(source: str, start: int, end: int) -> list[VbaToken]:
    """Memoized span-relative tokens INCLUDING comments and newlines.

    The raw view shares the derive machinery (and its per-source cache) with
    cached_statement_tokens under a distinct key, for callers that need to see
    comment placement (e.g. the Call-without-parens argument-list scan)."""
    entry: dict[tuple[int, int], list[VbaToken]] | None = None
    for i, (cached_source, by_span) in enumerate(_statement_token_cache):
        if cached_source == source:
            entry = by_span
            if i > 0:
                _statement_token_cache.insert(0, _statement_token_cache.pop(i))
            break
    if entry is None:
        entry = {}
        _statement_token_cache.insert(0, (source, entry))
        if len(_statement_token_cache) > _STATEMENT_TOKEN_CACHE_MAX:
            _statement_token_cache.pop()
    # Raw entries use a negated key space so they never collide with the
    # filtered entries for the same span ((start, end) with start >= 0).
    key = (-start - 1, -end - 1)
    toks = entry.get(key)
    if toks is None:
        toks = _derive_statement_tokens(source, start, end, keep_all=True)
        if toks is None:
            toks = list(tokenize(source[start:end]))
        entry[key] = toks
    return toks


def _derive_statement_tokens(
    source: str, start: int, end: int, keep_all: bool = False
) -> list[VbaToken] | None:
    """Derive a statement's span-relative significant tokens from the module's
    shared token stream instead of re-lexing the statement's text.

    The whole module is already tokenized once (tokenize_cached); re-running the
    lexer per statement was the analysis pass's largest remaining cost on big
    modules. Statement spans start at statement boundaries, which are line-start
    lexer contexts in both the module stream and an isolated slice, so the token
    streams agree; if a module token ever straddles the span boundary (which a
    well-formed statement span never produces), return None and the caller falls
    back to lexing the slice."""
    all_tokens = tokenize_cached(source)
    # Binary search: first token ending after the span starts.
    lo = 0
    hi = len(all_tokens) - 1
    first = len(all_tokens)
    while lo <= hi:
        mid = (lo + hi) >> 1
        if all_tokens[mid].end > start:
            first = mid
            hi = mid - 1
        else:
            lo = mid + 1
    out: list[VbaToken] = []
    for i in range(first, len(all_tokens)):
        tok = all_tokens[i]
        if tok.start >= end:
            break
        if tok.start < start or tok.end > end:
            return None
        if not keep_all and (tok.kind is TokenKind.COMMENT or tok.kind is TokenKind.NEWLINE):
            continue
        out.append(
            VbaToken(
                kind=tok.kind,
                raw_text=tok.raw_text,
                start=tok.start - start,
                end=tok.end - start,
                line=tok.line,
                character=tok.character,
                canonical_text=tok.canonical_text,
                leading_trivia=tok.leading_trivia,
                trailing_trivia=tok.trailing_trivia,
            )
        )
    return out


class _HasStart(Protocol):
    @property
    def start(self) -> int: ...


class _HasSpan(Protocol):
    @property
    def start(self) -> int: ...

    @property
    def end(self) -> int: ...


def statement_tokens_cached(source: str, span: _HasSpan) -> list[VbaToken]:
    """Significant tokens of a statement span, excluding comments and newlines
    (memoized per pass): upstream's statementTokensCached signature over
    cached_statement_tokens."""
    return cached_statement_tokens(source, span.start, span.end)


def first_token_at_or_after(tokens: Sequence[_HasStart], offset: int) -> int:
    """Lower bound in a source-ordered token stream; returns length past the last token."""
    lo = 0
    hi = len(tokens)
    while lo < hi:
        mid = lo + (hi - lo) // 2
        if tokens[mid].start < offset:
            lo = mid + 1
        else:
            hi = mid
    return lo


def _identifier_start(ch: str) -> bool:
    return ch == "_" or unicodedata.category(ch).startswith("L")


def _identifier_part(ch: str) -> bool:
    return ch == "_" or unicodedata.category(ch)[0] in ("L", "M", "N")


def is_identifier(text: str) -> bool:
    """True when the whole text is one identifier: a letter or underscore, then
    letters, marks, digits and underscores (XLIDE's IDENT_RE,
    `^[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*$`). VBA identifiers may use any locale's
    letters, and an ASCII-only test made `Dim g As Прибор` resolve to no type."""
    return bool(text) and _identifier_start(text[0]) and all(_identifier_part(ch) for ch in text[1:])


class _IdentPattern:
    """XLIDE's IDENT_RE, ``/^[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*$/u``. Python's ``re``
    has no Unicode property classes, so ``test`` checks by category."""

    __slots__ = ()

    def test(self, text: str) -> bool:
        return is_identifier(text)


IDENT_RE = _IdentPattern()


def identifiers_in(text: str) -> list[str]:
    """Every identifier-shaped word in free text: a letter or underscore, then
    letters, marks, digits and underscores (XLIDE's `[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*`).

    Python's `re` has no Unicode property classes, so this scans by category."""
    words: list[str] = []
    i = 0
    length = len(text)
    while i < length:
        if not _identifier_start(text[i]):
            i += 1
            continue
        start = i
        i += 1
        while i < length and _identifier_part(text[i]):
            i += 1
        words.append(text[start:i])
    return words


def identifier_words(text: str) -> list[str]:
    """The same words lower-cased, the way VBA compares names."""
    return [word.lower() for word in identifiers_in(text)]


def token_name(token: VbaToken | None) -> str | None:
    """Identifier-like name of a token (unwraps bracketed identifiers)."""
    if token is None:
        return None
    if token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD):
        return token.raw_text
    if token.kind is TokenKind.BRACKETED_IDENTIFIER:
        # Strip the surrounding brackets, but only when both are present: the tokenizer
        # still emits a bracketedIdentifier for an unterminated `[name` at line end,
        # where a blind slice would drop a real character.
        raw = token.raw_text
        return raw[1:-1] if raw.startswith("[") and raw.endswith("]") else raw
    return None


# A relational operator (MS-VBAL 5.6.9.5) in its standard spelling.
RelationalOperator = Literal["=", "<>", "<", ">", "<=", ">="]

_RELATIONAL_OPERATORS = frozenset({"=", "<>", "<", ">", "<=", ">="})

# Two single-character relational tokens that make one operator, in either order.
_RELATIONAL_PAIRS = {"<>": "<>", "><": "<>", "<=": "<=", "=<": "<=", ">=": ">=", "=>": ">="}


def relational_operator_at(
    tokens: Sequence[VbaToken], index: int, end: int | None = None
) -> tuple[str, int] | None:
    """The relational operator ``tokens[index]`` starts, in its standard spelling,
    and how many tokens it takes (1 or 2).

    MS-VBAL 5.6.9.5 writes `<>`, `<=` and `>=` as two special tokens in either
    order, so whitespace may stand between them: the VBE reads `a < > b` as
    `a <> b` and `a = > b` as `a >= b`. The lexer joins the two when they touch.
    Tokens at or past ``end`` are out of reach.
    """
    limit = len(tokens) if end is None else end
    token = tokens[index] if 0 <= index < limit else None
    if token is None or token.kind is not TokenKind.OPERATOR:
        return None
    text = token.canonical_text if token.canonical_text is not None else token.raw_text
    if text not in _RELATIONAL_OPERATORS:
        return None
    following = tokens[index + 1] if index + 1 < limit else None
    pair = (
        _RELATIONAL_PAIRS.get(token.raw_text + following.raw_text)
        if following is not None and following.kind is TokenKind.OPERATOR
        else None
    )
    return (pair, 2) if pair is not None else (text, 1)


def token_word(token: VbaToken | None) -> str:
    """Canonical (case-folded) text of a token, used for keyword matching.

    Keyword tokens carry canonical_text from the lexer; otherwise the raw text is
    lowered because VBA is case-insensitive (MS-VBAL 3.3.5).
    """
    if token is None:
        return ""
    text = token.canonical_text if token.canonical_text is not None else token.raw_text
    return text.lower()


def is_decimal_line_number(token: VbaToken | None) -> bool:
    """True when the token is a decimal line-number literal."""
    return (
        token is not None
        and token.kind is TokenKind.INTEGER_LITERAL
        and _DECIMAL_RE.match(token.raw_text) is not None
    )


def starts_physical_line(source: str, offset: int) -> bool:
    """True when only spaces and tabs stand between the line's start and
    ``offset``. A label is one only there, after the line number if there is
    one: in ``10: L1:`` the VBE reads L1 as a call (XLIDE issue #230)."""
    i = offset - 1
    while i >= 0 and source[i] in (" ", "\t"):
        i -= 1
    return i < 0 or source[i] in ("\n", "\r")


def tokens_without_leading_line_number(tokens: Sequence[VbaToken]) -> list[VbaToken]:
    """Drops the leading line-number token when one prefixes the statement."""
    if len(tokens) > 1 and is_decimal_line_number(tokens[0]):
        return list(tokens[1:])
    return list(tokens)


def match_paren_from(tokens: Sequence[VbaToken], open_index: int) -> int:
    """Index of the ')' matching the '(' at open_index, or -1 when unmatched."""
    depth = 0
    for i in range(open_index, len(tokens)):
        raw = tokens[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


def split_top_level_token_groups(
    tokens: Sequence[VbaToken],
    start: int,
    separator: str,
    end: int | None = None,
) -> list[list[VbaToken]]:
    """Split tokens[start:end) into separator-delimited groups at paren depth 0."""
    if end is None:
        end = len(tokens)
    groups: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for i in range(start, end):
        raw = tokens[i].raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth = max(0, depth - 1)
        if depth == 0 and raw == separator:
            groups.append(current)
            current = []
            continue
        current.append(tokens[i])
    groups.append(current)
    return groups


def top_level_equals_index(tokens: Sequence[VbaToken]) -> int:
    """First '=' outside parentheses, retaining raw-token assignment semantics."""
    depth = 0
    for i, token in enumerate(tokens):
        raw = token.raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif raw == "=" and depth == 0:
            return i
    return -1


_REM_COMMENT_RE = re.compile(r"^rem\b", re.IGNORECASE | re.ASCII)
_REM_RE = re.compile(r"rem", re.IGNORECASE)


def directive_rem(tokens: Sequence[VbaToken], hash_index: int) -> VbaToken | None:
    """The `Rem` on the compiler-directive statement whose `#` is
    ``tokens[hash_index]``: the word itself (it stays one after `Then`) or the
    comment the lexer starts with it, but not a member name (`x.Rem`). The VBE
    does not read Rem as a comment on a directive line: an #If, #ElseIf, #Else
    or #End If line holding one is no directive at all, and a #Const line
    holding one is a syntax error (XLIDE issue #1321). The scan stops at the
    statement's newline or colon, and an apostrophe comment ends it."""
    for i in range(hash_index + 1, len(tokens)):
        tok = tokens[i]
        if tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.COLON:
            return None
        if tok.kind is TokenKind.COMMENT:
            return tok if _REM_COMMENT_RE.match(tok.raw_text) else None
        # `x.Rem` and `x!Rem` name a member: that line is still a directive.
        if token_word(tok) == "rem" and tokens[i - 1].raw_text not in (".", "!"):
            return tok
    return None


_DIRECTIVE_WORDS = frozenset({"if", "elseif", "else", "end", "endif", "const"})
_NO_STARTS: frozenset[int] = frozenset()
# One scan per token array: the parser and malformed-statement each ask, and
# malformed-statement asks once per line.
_rem_directive_starts_by_tokens = IdentityLru()


def rem_directive_starts(source: str) -> frozenset[int]:
    """The start offsets of the `#` of every #If, #ElseIf, #Else, #End If or
    #Const line holding a `Rem`, which the VBE does not act on as that
    directive (XLIDE issue #1321): a #Const one defines nothing, even in a
    skipped arm, where a plain #Const still does. Only a `#` that starts its
    physical line counts (the token before it is a newline: a continued line
    has none)."""
    if _REM_RE.search(source) is None:
        return _NO_STARTS
    tokens = tokenize_cached(source)
    cached: frozenset[int] | None = _rem_directive_starts_by_tokens.get(tokens)
    if cached is not None:
        return cached
    starts = frozenset(
        tokens[i].start
        for i in range(len(tokens))
        if tokens[i].kind is TokenKind.DIRECTIVE
        and (i == 0 or tokens[i - 1].kind is TokenKind.NEWLINE)
        and token_word(tokens[i + 1] if i + 1 < len(tokens) else None) in _DIRECTIVE_WORDS
        and directive_rem(tokens, i) is not None
    )
    _rem_directive_starts_by_tokens.put(starts, tokens)
    return starts


def directive_colon(tokens: Sequence[VbaToken], hash_index: int) -> VbaToken | None:
    """The statement colon after the compiler directive whose `#` is
    ``tokens[hash_index]``, on its own line."""
    for i in range(hash_index + 1, len(tokens)):
        tok = tokens[i]
        if tok.kind is TokenKind.COLON:
            return tok
        if tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.COMMENT:
            return None
    return None


_colon_directive_starts_by_tokens = IdentityLru()


def colon_directive_starts(source: str) -> frozenset[int]:
    """The start offsets of the `#` of every #If, #ElseIf, #Else, #End If or
    #Const line a statement colon follows, with nothing, a comment or code
    after it. The VBE does not read such a line as a directive at all: it
    opens, divides, closes and defines nothing, is skipped with a skipped arm,
    and is an error where it is compiled (XLIDE issue #1323, measured in Excel
    16.0). As for a Rem, only a `#` that starts its physical line counts."""
    if ":" not in source:
        return _NO_STARTS
    tokens = tokenize_cached(source)
    cached: frozenset[int] | None = _colon_directive_starts_by_tokens.get(tokens)
    if cached is not None:
        return cached
    starts = frozenset(
        tokens[i].start
        for i in range(len(tokens))
        if tokens[i].kind is TokenKind.DIRECTIVE
        and (i == 0 or tokens[i - 1].kind is TokenKind.NEWLINE)
        and token_word(tokens[i + 1] if i + 1 < len(tokens) else None) in _DIRECTIVE_WORDS
        and directive_colon(tokens, i) is not None
    )
    _colon_directive_starts_by_tokens.put(starts, tokens)
    return starts


_misplaced_directive_starts_by_tokens = IdentityLru()


def misplaced_directive_starts(source: str) -> frozenset[int]:
    """The start offsets of the `#` of every #If, #ElseIf, #Else, #End If or
    #Const that does not start its logical line: after a statement's or a
    label's colon, another directive's colon, or a line number. The VBE reads
    a directive only where it starts its logical line, after nothing but
    blanks and line continuations; elsewhere it opens, divides, closes and
    defines nothing, and is an error where it is compiled (XLIDE issue #1331,
    measured in Excel 16.0). The lexer makes a directive token after a colon,
    and an operator after a line number; both are listed."""
    if "#" not in source:
        return _NO_STARTS
    tokens = tokenize_cached(source)
    cached: frozenset[int] | None = _misplaced_directive_starts_by_tokens.get(tokens)
    if cached is not None:
        return cached
    found: list[int] = []
    for i in range(1, len(tokens)):
        tok = tokens[i]
        following = tokens[i + 1] if i + 1 < len(tokens) else None
        if tok.raw_text != "#" or token_word(following) not in _DIRECTIVE_WORDS:
            continue
        after_colon = tok.kind is TokenKind.DIRECTIVE and tokens[i - 1].kind is not TokenKind.NEWLINE
        after_line_number = (
            tok.kind is TokenKind.OPERATOR
            and is_decimal_line_number(tokens[i - 1])
            and (i == 1 or tokens[i - 2].kind is TokenKind.NEWLINE)
        )
        if after_colon or after_line_number:
            found.append(tok.start)
    starts = frozenset(found)
    _misplaced_directive_starts_by_tokens.put(starts, tokens)
    return starts
