"""M1: lexer-derived stripped-line substrate (strippedLines.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.lexer.stripped_lines import lexer_stripped_line, lexer_stripped_lines


def test_blanks_string_literal_preserving_length() -> None:
    line = 's = "hello"'  # the literal "hello" occupies columns 4..10
    stripped = lexer_stripped_line(line)
    assert len(stripped) == len(line)
    assert stripped[:4] == "s = "
    assert stripped[4:] == " " * 7
    assert '"' not in stripped and "hello" not in stripped


def test_blanks_trailing_comment() -> None:
    line = "x = 1 ' secret"
    stripped = lexer_stripped_line(line)
    assert len(stripped) == len(line)
    assert stripped[:6] == "x = 1 "
    assert stripped.rstrip() == "x = 1"
    assert "secret" not in stripped


def test_code_outside_strings_and_comments_is_untouched() -> None:
    line = "Call Foo(1, 2)"
    assert lexer_stripped_line(line) == line


def test_multiline_blanks_each_physical_line_independently() -> None:
    source = 'a = "x"\nb = 2 ' + "' note"
    lines = lexer_stripped_lines(source)
    src_lines = source.split("\n")
    assert len(lines) == 2
    assert all(len(out) == len(src) for out, src in zip(lines, src_lines))
    assert lines[0].rstrip() == "a ="
    assert lines[1].rstrip() == "b = 2"
    assert '"' not in lines[0] and "note" not in lines[1]


def test_rem_comment_after_colon_is_blanked() -> None:
    # The lexer recognizes Rem at any statement start (MS-VBAL 3.3.5.2), so a
    # trailing ": Rem ..." comment is blanked (the deliberate divergence the
    # source documents versus the legacy whole-line-only stripVba scanner).
    line = "x = 1: Rem trailing"
    stripped = lexer_stripped_line(line)
    assert len(stripped) == len(line)
    assert stripped.startswith("x = 1:")
    assert "Rem" not in stripped and "trailing" not in stripped
