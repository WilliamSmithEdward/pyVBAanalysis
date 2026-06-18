"""M1: lexer round-trip over the oracle corpus plus focused token tests."""

from __future__ import annotations

import pytest

from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.lexer import TokenKind, VbaToken, tokenize

_SOURCES: list[tuple[str, str]] = [
    (f"{c.id}::{m.name}", m.source) for c in load_oracle_cases() for m in c.modules
]


def _reconstruct(tokens: list[VbaToken]) -> str:
    parts: list[str] = []
    for tok in tokens:
        for tr in tok.leading_trivia:
            parts.append(tr.text)
        parts.append(tok.raw_text)
    if tokens:
        for tr in tokens[-1].trailing_trivia:
            parts.append(tr.text)
    return "".join(parts)


@pytest.mark.parametrize("source", [s for _, s in _SOURCES], ids=[i for i, _ in _SOURCES])
def test_roundtrip_on_oracle_sources(source: str) -> None:
    # The acceptance gate: re-joining trivia + raw text reproduces the source.
    assert _reconstruct(tokenize(source)) == source


def test_keyword_vs_identifier() -> None:
    kinds = {t.raw_text: t.kind for t in tokenize("Dim x As Long")}
    assert kinds["Dim"] == TokenKind.KEYWORD
    assert kinds["As"] == TokenKind.KEYWORD
    assert kinds["Long"] == TokenKind.KEYWORD
    assert kinds["x"] == TokenKind.IDENTIFIER


def test_canonical_casing() -> None:
    (tok,) = [t for t in tokenize("diM") if t.kind == TokenKind.KEYWORD]
    assert tok.canonical_text == "Dim"


def test_rem_comment_at_statement_start() -> None:
    toks = tokenize("Rem a comment")
    assert toks[0].kind == TokenKind.COMMENT
    assert toks[0].raw_text == "Rem a comment"


def test_apostrophe_comment() -> None:
    comments = [t for t in tokenize("x = 1 ' note") if t.kind == TokenKind.COMMENT]
    assert comments and comments[0].raw_text == "' note"


def test_number_literals() -> None:
    assert tokenize("&HFF")[0].kind == TokenKind.INTEGER_LITERAL
    assert tokenize("&O17")[0].kind == TokenKind.INTEGER_LITERAL
    assert tokenize("100&")[0].kind == TokenKind.INTEGER_LITERAL
    assert tokenize("1.5")[0].kind == TokenKind.FLOAT_LITERAL
    assert tokenize("1.5e3")[0].kind == TokenKind.FLOAT_LITERAL
    assert tokenize("3.14#")[0].kind == TokenKind.FLOAT_LITERAL


def test_member_access_dot_not_part_of_number() -> None:
    # `a.b` - the '.' is punctuation, not a decimal point.
    kinds = [t.kind for t in tokenize("a.b")]
    assert kinds == [TokenKind.IDENTIFIER, TokenKind.PUNCTUATION, TokenKind.IDENTIFIER]


def test_date_literal_in_expression() -> None:
    dates = [t for t in tokenize("d = #1/1/2020#") if t.kind == TokenKind.DATE_LITERAL]
    assert dates and dates[0].raw_text == "#1/1/2020#"


def test_hash_file_number_is_operator() -> None:
    # `Write #1, x` - the '#' is a file-number marker, not a date literal.
    hashes = [t for t in tokenize("Write #1, x") if t.raw_text == "#"]
    assert hashes and hashes[0].kind == TokenKind.OPERATOR


def test_directive_hash_at_statement_start() -> None:
    assert tokenize("#If VBA7 Then")[0].kind == TokenKind.DIRECTIVE


def test_multi_char_operators() -> None:
    assert [t.raw_text for t in tokenize("a <= b") if t.kind == TokenKind.OPERATOR] == ["<="]
    assert [t.raw_text for t in tokenize("a <> b") if t.kind == TokenKind.OPERATOR] == ["<>"]
    assert [t.raw_text for t in tokenize("x := 1") if t.kind == TokenKind.OPERATOR] == [":="]


def test_line_continuation_roundtrips() -> None:
    src = "x = 1 + _\n    2"
    assert _reconstruct(tokenize(src)) == src


def test_unterminated_string_ends_at_line_end() -> None:
    toks = tokenize('s = "abc\nx = 1')
    strings = [t for t in toks if t.kind == TokenKind.STRING_LITERAL]
    assert strings and strings[0].raw_text == '"abc'
