"""M1: statement-level token helpers (tokenHelpers.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.lexer import tokenize
from pyvbaanalysis.lexer.token_helpers import (
    is_decimal_line_number,
    is_ident_like,
    match_paren_from,
    split_top_level_token_groups,
    statement_tokens,
    token_name,
    token_word,
    tokens_without_leading_line_number,
)


def test_is_ident_like() -> None:
    toks = {t.raw_text: t for t in tokenize("Dim x123 As Long")}
    assert is_ident_like(toks["Dim"])  # keyword reads as identifier
    assert is_ident_like(toks["x123"])
    op = next(t for t in tokenize("a + b") if t.raw_text == "+")
    assert not is_ident_like(op)


def test_statement_tokens_drops_comments_and_newlines() -> None:
    source = "x = 1 ' note\ny = 2"
    kinds = [t.raw_text for t in statement_tokens(source, 0, len(source))]
    assert "'" not in "".join(kinds)
    assert "note" not in kinds
    assert kinds == ["x", "=", "1", "y", "=", "2"]


def test_statement_tokens_respects_span() -> None:
    source = "Dim a: Dim b"
    first = statement_tokens(source, 0, 5)  # "Dim a"
    assert [t.raw_text for t in first] == ["Dim", "a"]


def test_token_name_plain_and_bracketed() -> None:
    assert token_name(tokenize("Range")[0]) == "Range"
    assert token_name(tokenize("[My Sheet!]")[0]) == "My Sheet!"
    assert token_name(tokenize("+")[0]) is None
    assert token_name(None) is None


def test_token_word_uses_canonical_casing() -> None:
    (kw,) = [t for t in tokenize("dIm") if t.canonical_text is not None]
    assert token_word(kw) == "dim"
    assert token_word(tokenize("Foo")[0]) == "foo"
    assert token_word(None) == ""


def test_is_decimal_line_number() -> None:
    assert is_decimal_line_number(tokenize("100")[0])
    assert not is_decimal_line_number(tokenize("&HFF")[0])
    assert not is_decimal_line_number(tokenize("1.5")[0])
    assert not is_decimal_line_number(None)


def test_tokens_without_leading_line_number() -> None:
    toks = statement_tokens("100 GoTo Done", 0, len("100 GoTo Done"))
    stripped = tokens_without_leading_line_number(toks)
    assert [t.raw_text for t in stripped] == ["GoTo", "Done"]
    # A lone line number is preserved (length-1 guard).
    only = statement_tokens("100", 0, 3)
    assert [t.raw_text for t in tokens_without_leading_line_number(only)] == ["100"]


def test_match_paren_from() -> None:
    toks = statement_tokens("Foo(a, (b))", 0, len("Foo(a, (b))"))
    open_index = next(i for i, t in enumerate(toks) if t.raw_text == "(")
    close = match_paren_from(toks, open_index)
    assert toks[close].raw_text == ")"
    assert close == len(toks) - 1  # the outer ')'
    # Unmatched returns -1.
    bad = statement_tokens("Foo(a", 0, len("Foo(a"))
    bad_open = next(i for i, t in enumerate(bad) if t.raw_text == "(")
    assert match_paren_from(bad, bad_open) == -1


def test_split_top_level_token_groups() -> None:
    toks = statement_tokens("a, f(b, c), d", 0, len("a, f(b, c), d"))
    groups = split_top_level_token_groups(toks, 0, ",")
    rendered = [[t.raw_text for t in g] for g in groups]
    assert rendered == [["a"], ["f", "(", "b", ",", "c", ")"], ["d"]]
