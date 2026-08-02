"""Small token utilities shared across analyzer surfaces.

Ported from xlide_vscode/src/analyzer/lexer/tokenHelpers.ts. Keeps statement-level
token handling (comment/newline filtering, identifier extraction, leading line
numbers, paren matching) from drifting between surfaces. VBA is case-insensitive
(MS-VBAL 3.3.5), so name matching folds case.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from .token_kinds import TokenKind, VbaToken
from .tokenize import tokenize, tokenize_cached

IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_DECIMAL_RE = re.compile(r"^\d+$")


def is_ident_like(token: VbaToken) -> bool:
    """True when the token reads as a bare identifier (identifier or keyword)."""
    return (
        token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
        and IDENT_RE.match(token.raw_text) is not None
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
