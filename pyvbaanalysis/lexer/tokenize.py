"""VBA tokenizer.

Ported from xlide_vscode/src/analyzer/lexer/tokenize.ts. Verified against MS-VBAL
v20250520 (3.2.2 logical lines, 3.3.1 separators/special tokens, 3.3.2 numbers,
3.3.3 dates, 3.3.4 strings, 3.3.5 identifiers, 3.4 conditional compilation).

The token stream is loss-aware and round-trippable: re-joining every token's
leading trivia and raw_text (and any trailing trivia on the final token)
reproduces the source exactly. Reserved keywords carry canonical capitalization,
and so do contextual keywords inside the statement that makes them keywords.
"""

from __future__ import annotations

import re
import unicodedata
from functools import lru_cache

from .contextual_keywords import settle_contextual_keywords
from .keyword_table import canonical_keyword
from .token_kinds import TokenKind, VbaToken, is_line_terminator, is_wsc
from .trivia import continuation_terminator, scan_leading_trivia

# Type-suffix chars (MS-VBAL 3.3.2).
_INTEGER_SUFFIX = frozenset(("%", "&", "^"))
_FLOAT_SUFFIX = frozenset(("!", "#", "@"))


def _is_digit(ch: str) -> bool:
    return "0" <= ch <= "9"


def _is_octal_digit(ch: str) -> bool:
    return "0" <= ch <= "7"


def _is_hex_digit(ch: str) -> bool:
    return _is_digit(ch) or ("a" <= ch <= "f") or ("A" <= ch <= "F")


def _is_exponent_letter(ch: str) -> bool:
    return ch in ("e", "E", "d", "D")


def _is_ident_start(ch: str) -> bool:
    if ("A" <= ch <= "Z") or ("a" <= ch <= "z"):
        return True
    # Non-Latin identifier forms (MS-VBAL 3.3.5.1): permit Unicode letters.
    return ord(ch) >= 0x80 and ch.isalpha()


def _is_ident_part(ch: str) -> bool:
    if ch == "_" or _is_digit(ch):
        return True
    if _is_ident_start(ch):
        return True
    # A combining mark continues the identifier it is attached to. Scripts like
    # Thai write a single letter as a base plus a tone mark and/or vowel sign,
    # and str.isalpha() is False for those marks (categories Mn and Mc), so
    # without this the identifier splits mid-grapheme and the tail is reported
    # as an undeclared variable.
    #
    # VBE-oracle verified 2026-08-03: a cp874 project declaring
    # `Dim <kho + mai ek + sara aa> As String` compiles and runs clean in real
    # Excel (no compile-error dialog), alongside a mark-free Thai control that
    # also runs. The identifier is legal, so splitting it was a false positive.
    # Every mark category (Mn, Mc, Me), matching XLIDE's \p{M}: the port follows
    # upstream as the executable spec, and widening what continues a name can
    # only remove false positives, never add one.
    return ord(ch) >= 0x80 and unicodedata.category(ch).startswith("M")


# The ASCII subset of _is_ident_part, as a run matcher. Always matches (possibly
# empty), so `.end()` is the scan position after the run.
_ASCII_IDENT_RUN_RE = re.compile(r"[A-Za-z0-9_]*")


def tokenize(src: str) -> list[VbaToken]:
    """Tokenize a VBA module body into a flat, round-trippable token stream."""
    tokens: list[VbaToken] = []
    length = len(src)
    pos = 0
    line = 0
    character = 0
    # EOS / statement-start tracking: true at file start and after a newline or a
    # ':' separator (MS-VBAL 3.3.1 EOS). Governs Rem-comment and directive lexing.
    at_statement_start = True

    while pos < length:
        scan = scan_leading_trivia(src, pos, line, character)
        leading = scan.trivia
        pos, line, character = scan.pos, scan.line, scan.character

        if pos >= length:
            # Trailing trivia at EOF: attach to the last token so the stream stays
            # round-trippable. With no tokens, the input was trivia-only.
            if leading and tokens:
                tokens[-1].trailing_trivia = tuple(leading)
            break

        start_pos = pos
        start_line = line
        start_char = character
        ch = src[pos]
        kind: TokenKind
        canonical: str | None = None
        is_newline = False

        if is_line_terminator(ch):
            if ch == "\r" and pos + 1 < length and src[pos + 1] == "\n":
                pos += 2
            else:
                pos += 1
            kind = TokenKind.NEWLINE
            is_newline = True
            at_statement_start = True
        elif ch == "'":
            # Apostrophe comment to the end of the logical line (MS-VBAL 3.3.1).
            pos = _comment_end(src, pos + 1)
            kind = TokenKind.COMMENT
        elif _is_ident_start(ch):
            # Scan the identifier in ASCII runs (one regex step each) instead of
            # one predicate call per character; a non-ASCII continuation char is
            # still checked with the exact predicate and then the run resumes, so
            # the stopping point matches the per-character walk exactly.
            pos = _ASCII_IDENT_RUN_RE.match(src, pos + 1).end()  # type: ignore[union-attr]
            while pos < length and _is_ident_part(src[pos]):
                pos = _ASCII_IDENT_RUN_RE.match(src, pos + 1).end()  # type: ignore[union-attr]
            word = src[start_pos:pos]
            if word.lower() == "rem" and at_statement_start:
                # Rem comment (MS-VBAL 3.3.5.2): rest of line is comment.
                pos = _comment_end(src, pos)
                kind = TokenKind.COMMENT
            else:
                canonical = canonical_keyword(word)
                kind = TokenKind.KEYWORD if canonical else TokenKind.IDENTIFIER
                at_statement_start = False
        elif (
            _is_digit(ch)
            or (ch == "." and pos + 1 < length and _is_digit(src[pos + 1]))
            or (
                ch == "&"
                and pos + 1 < length
                and (src[pos + 1] in ("h", "H", "o", "O") or _is_octal_digit(src[pos + 1]))
            )
        ):
            # The O of an octal literal is optional (MS-VBAL 3.3.2), and the VBE
            # takes `&17` as one wherever it stands: `x = "a" &1` does not
            # compile, while `x = "a" &9` is a concatenation (XLIDE issue #87).
            kind, pos = _lex_number(src, pos)
            at_statement_start = False
        elif ch == '"':
            # String literal (MS-VBAL 3.3.4): doubled-quote escaping; may end at the
            # closing quote or at LINE-END (unterminated tolerated).
            pos += 1
            while pos < length:
                c = src[pos]
                if c == '"':
                    if pos + 1 < length and src[pos + 1] == '"':
                        pos += 2  # escaped quote
                        continue
                    pos += 1  # closing quote
                    break
                if is_line_terminator(c):
                    break  # unterminated string ends at LINE-END
                pos += 1
            kind = TokenKind.STRING_LITERAL
            at_statement_start = False
        elif ch == "[":
            # FOREIGN-NAME (MS-VBAL 3.3.5.3): "[" 1*non-line-termination-char "]".
            pos += 1
            while pos < length and src[pos] != "]" and not is_line_terminator(src[pos]):
                pos += 1
            if pos < length and src[pos] == "]":
                pos += 1
            kind = TokenKind.BRACKETED_IDENTIFIER
            at_statement_start = False
        elif ch == "#":
            if at_statement_start:
                # Conditional-compilation directive marker (MS-VBAL 3.4).
                pos += 1
                kind = TokenKind.DIRECTIVE
                at_statement_start = False
            else:
                # Candidate '#'-delimited date literal on the same physical line. The
                # '#' pair only forms a DATE token when the enclosed text is a valid
                # date-or-time body (MS-VBAL 3.3.3); otherwise this '#' is a
                # file-number marker or stray type-suffix and lexes as an operator.
                scan_p = pos + 1
                while scan_p < length and src[scan_p] != "#" and not is_line_terminator(src[scan_p]):
                    scan_p += 1
                if (
                    scan_p < length
                    and src[scan_p] == "#"
                    and _is_date_literal_body(src[pos + 1 : scan_p])
                ):
                    pos = scan_p + 1
                    kind = TokenKind.DATE_LITERAL
                else:
                    pos += 1
                    kind = TokenKind.OPERATOR
                at_statement_start = False
        else:
            kind, pos = _lex_symbol(src, ch, pos)
            canonical = _REVERSED_RELATIONAL_OPERATORS.get(src[start_pos:pos])
            at_statement_start = kind == TokenKind.COLON

        token = VbaToken(
            kind=kind,
            raw_text=src[start_pos:pos],
            start=start_pos,
            end=pos,
            line=start_line,
            character=start_char,
            canonical_text=canonical,
            leading_trivia=tuple(leading),
        )
        tokens.append(token)

        if is_newline:
            line += 1
            character = 0
        elif kind is TokenKind.COMMENT:
            # A comment may run on through line continuations.
            raw = token.raw_text
            last_break = max(raw.rfind("\n"), raw.rfind("\r"))
            if last_break < 0:
                character += pos - start_pos
            else:
                line += _line_break_count(raw)
                character = len(raw) - last_break - 1
        else:
            character += pos - start_pos

    settle_contextual_keywords(tokens)
    return tokens


def _comment_end(src: str, start: int) -> int:
    """Where a comment whose body starts at ``start`` ends. A comment-body runs
    through line-continuations to LINE-END (MS-VBAL 3.3.1), so the VBE takes a
    comment ending in ` _` on through the next line (XLIDE issue #82)."""
    length = len(src)
    pos = start
    while pos < length and not is_line_terminator(src[pos]):
        terminator = (
            continuation_terminator(src, pos) if src[pos] == "_" and pos > 0 and is_wsc(src[pos - 1]) else -1
        )
        if terminator >= 0:
            step = 2 if src[terminator] == "\r" and terminator + 1 < length and src[terminator + 1] == "\n" else 1
            pos = terminator + step
            continue
        pos += 1
    return pos


def _line_break_count(text: str) -> int:
    """Line breaks in ``text``, a CRLF counting once."""
    count = 0
    for i, ch in enumerate(text):
        if ch == "\n" or (ch == "\r" and (i + 1 >= len(text) or text[i + 1] != "\n")):
            count += 1
    return count


@lru_cache(maxsize=8)
def tokenize_cached(src: str) -> tuple[VbaToken, ...]:
    """Read-only memoized tokenize for hot paths. Do not mutate the result."""
    return tuple(tokenize(src))


def _lex_number(src: str, p: int) -> tuple[TokenKind, int]:
    """Lex a numeric literal at p; return (kind, new_pos). MS-VBAL 3.3.2."""
    length = len(src)
    ch = src[p]

    if ch == "&":
        # Hex (&H) or octal (&O, or & with the O left out) integer literal.
        radix = src[p + 1]
        p += 1 if _is_octal_digit(radix) else 2  # consume '&' and any radix letter
        if radix in ("h", "H"):
            while p < length and _is_hex_digit(src[p]):
                p += 1
        else:
            while p < length and _is_octal_digit(src[p]):
                p += 1
        if p < length and src[p] in _INTEGER_SUFFIX:
            p += 1
        return TokenKind.INTEGER_LITERAL, p

    is_float = False
    while p < length and _is_digit(src[p]):
        p += 1
    # Optional decimal point, with the fractional digits optional too (MS-VBAL
    # 3.3.2): the VBE reads `1.` as `1#` (XLIDE issue #87). A letter after the dot
    # that does not start an exponent leaves it a member-access dot.
    if p < length and src[p] == ".":
        after = src[p + 1] if p + 1 < length else ""
        if (
            _is_digit(after)
            or (_is_exponent_letter(after) and _has_exponent_tail(src, p + 1))
            or after == ""
            or not _is_ident_start(after)
        ):
            is_float = True
            p += 1
            while p < length and _is_digit(src[p]):
                p += 1
    # Optional exponent.
    if p < length and _is_exponent_letter(src[p]) and _has_exponent_tail(src, p):
        is_float = True
        p += 1
        if p < length and src[p] in ("+", "-"):
            p += 1
        while p < length and _is_digit(src[p]):
            p += 1
    # Type suffix.
    if p < length:
        if src[p] in _FLOAT_SUFFIX:
            is_float = True
            p += 1
        elif not is_float and src[p] in _INTEGER_SUFFIX:
            p += 1
    return (TokenKind.FLOAT_LITERAL if is_float else TokenKind.INTEGER_LITERAL), p


def _has_exponent_tail(src: str, pos: int) -> bool:
    """True if pos begins a valid exponent tail: [DdEe] [sign] 1*digit."""
    length = len(src)
    if pos >= length or not _is_exponent_letter(src[pos]):
        return False
    p = pos + 1
    if p < length and src[p] in ("+", "-"):
        p += 1
    return p < length and _is_digit(src[p])


# Date-literal body validation (MS-VBAL 3.3.3). Whitespace alone is a
# date-separator, so the grammar is ambiguous; the matchers return every candidate
# end position and the body is accepted if any reading consumes it exactly.

_MONTH_NAMES = frozenset(
    (
        "january", "february", "march", "april", "may", "june", "july", "august",
        "september", "october", "november", "december",
        "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
    )
)


def _is_date_literal_body(body: str) -> bool:
    """True when the text between a '#' pair is a valid date-literal body."""
    start = _skip_wsc(body, 0)
    if start == len(body):
        return True  # "#" *WSC "#": the empty date literal
    for end in _date_or_time_ends(body, start):
        if _skip_wsc(body, end) == len(body):
            return True
    return False


def _date_or_time_ends(s: str, pos: int) -> list[int]:
    ends: list[int] = []
    for date_end in _date_value_ends(s, pos):
        ends.append(date_end)  # date-value alone
        ws_end = _skip_wsc(s, date_end)
        if ws_end > date_end:
            ends.extend(_time_value_ends(s, ws_end))  # date-value 1*WSC time-value
    ends.extend(_time_value_ends(s, pos))  # time-value alone
    return ends


def _date_value_ends(s: str, pos: int) -> list[int]:
    ends: list[int] = []
    left = _date_part_end(s, pos)
    if left < 0:
        return ends
    sep1 = _date_separator_end(s, left)
    if sep1 < 0:
        return ends
    middle = _date_part_end(s, sep1)
    if middle < 0:
        return ends
    ends.append(middle)
    sep2 = _date_separator_end(s, middle)
    if sep2 >= 0:
        right = _date_part_end(s, sep2)
        if right >= 0:
            ends.append(right)
    return ends


def _date_part_end(s: str, pos: int) -> int:
    digits_end = _decimal_end(s, pos)
    if digits_end >= 0:
        return digits_end
    p = pos
    while p < len(s) and (("A" <= s[p] <= "Z") or ("a" <= s[p] <= "z")):
        p += 1
    return p if p > pos and s[pos:p].lower() in _MONTH_NAMES else -1


def _date_separator_end(s: str, pos: int) -> int:
    after_ws = _skip_wsc(s, pos)
    ch = s[after_ws] if after_ws < len(s) else ""
    if ch in ("/", "-", ","):
        return _skip_wsc(s, after_ws + 1)
    return after_ws if after_ws > pos else -1


def _time_value_ends(s: str, pos: int) -> list[int]:
    ends: list[int] = []
    hour = _decimal_end(s, pos)
    if hour < 0:
        return ends
    hour_ampm = _ampm_end(s, hour)
    if hour_ampm >= 0:
        ends.append(hour_ampm)  # hour-value ampm
    sep1 = _time_separator_end(s, hour)
    if sep1 < 0:
        return ends
    minute = _decimal_end(s, sep1)
    if minute < 0:
        return ends
    ends.append(minute)
    minute_ampm = _ampm_end(s, minute)
    if minute_ampm >= 0:
        ends.append(minute_ampm)
    sep2 = _time_separator_end(s, minute)
    if sep2 >= 0:
        second = _decimal_end(s, sep2)
        if second >= 0:
            ends.append(second)
            second_ampm = _ampm_end(s, second)
            if second_ampm >= 0:
                ends.append(second_ampm)
    return ends


def _time_separator_end(s: str, pos: int) -> int:
    after_ws = _skip_wsc(s, pos)
    ch = s[after_ws] if after_ws < len(s) else ""
    return _skip_wsc(s, after_ws + 1) if ch in (":", ".") else -1


def _ampm_end(s: str, pos: int) -> int:
    p = _skip_wsc(s, pos)
    first = s[p].lower() if p < len(s) else ""
    if first not in ("a", "p"):
        return -1
    second = s[p + 1].lower() if p + 1 < len(s) else ""
    return p + 2 if second == "m" else p + 1


def _decimal_end(s: str, pos: int) -> int:
    p = pos
    while p < len(s) and _is_digit(s[p]):
        p += 1
    return p if p > pos else -1


def _skip_wsc(s: str, pos: int) -> int:
    p = pos
    while p < len(s) and is_wsc(s[p]):
        p += 1
    return p


# The standard spelling of a relational operator written the other way round.
# MS-VBAL 5.6.9.5 writes `<>`, `<=` and `>=` as two special tokens in either
# order, and the VBE stores `=>` as `>=` (XLIDE issue #87).
_REVERSED_RELATIONAL_OPERATORS: dict[str, str] = {"=>": ">=", "=<": "<=", "><": "<>"}


def _lex_symbol(src: str, ch: str, p: int) -> tuple[TokenKind, int]:
    """Lex an operator/punctuation/colon at p; return (kind, new_pos)."""
    length = len(src)
    nxt = src[p + 1] if p + 1 < length else ""

    # Multi-character operators :=, <=, >=, <>, and the reversed =>, =< and ><.
    if (
        (ch == ":" and nxt == "=")
        or (ch == "<" and nxt in ("=", ">"))
        or (ch == ">" and nxt == "=")
        or ch + nxt in _REVERSED_RELATIONAL_OPERATORS
    ):
        return TokenKind.OPERATOR, p + 2

    p += 1
    if ch == ":":
        return TokenKind.COLON, p
    if ch in (",", ".", "(", ")", ";"):
        return TokenKind.PUNCTUATION, p
    if ch in ("=", "<", ">", "+", "-", "*", "/", "\\", "^", "&", "!", "?"):
        return TokenKind.OPERATOR, p
    return TokenKind.UNKNOWN, p
