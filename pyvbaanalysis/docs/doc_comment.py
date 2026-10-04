"""The `'''` documentation-comment grammar: which lines belong to a declaration.

Ported from xlide_vscode/src/analyzer/docs/docComment.ts, the parts the diagnostics
need. The editor's hover rendering of a parsed doc is out of scope for this port.
"""

from __future__ import annotations

import re
from bisect import bisect_right
from collections.abc import Iterator
from dataclasses import dataclass

from ..js_compat import JS_WHITESPACE, js_trim
from ..parser.nodes import Span

_XLIDE_DIRECTIVE_RE = re.compile(r"^'+\s*@xlide-\S", re.IGNORECASE)
# ASCII: JavaScript's `\b` and case-insensitive matching are ASCII-only.
_HAS_TAG_RE = re.compile(
    r"<(summary|param|returns|remarks|example|signature)\b", re.IGNORECASE | re.ASCII
)
_JS_S = "[" + JS_WHITESPACE + "]"
_ATTRIBUTE_NAME_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_-]*")
_ATTRIBUTE_VALUE_OPENING_RE = re.compile(_JS_S + "*=" + _JS_S + '*"')
_WHITESPACE_RUN_RE = re.compile(_JS_S + "+")
_LEADING_WHITESPACE_RE = re.compile(r"^[ \t]*")


def detect_eol(source: str) -> str:
    """The module's line ending, so an inserted line matches its neighbours."""
    return "\r\n" if "\r\n" in source else "\n"


def leading_whitespace(text: str) -> str:
    match = _LEADING_WHITESPACE_RE.match(text)
    return match.group(0) if match is not None else ""


def _decode_entities(text: str) -> str:
    return (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def _collapse(text: str) -> str:
    """Trims and collapses internal whitespace runs (including newlines) to a space."""
    return js_trim(_WHITESPACE_RUN_RE.sub(" ", _decode_entities(text)))


@dataclass(frozen=True, slots=True)
class _DocAttribute:
    # Lowercased.
    name: str
    # Entity-decoded and trimmed.
    value: str
    # Raw value coordinates relative to the attribute string.
    start: int
    length: int


def _doc_attributes(raw: str) -> Iterator[_DocAttribute]:
    """The `name="value"` attributes of a tag's attribute text, advancing over
    attribute names without retrying every suffix of an unknown word."""
    at = 0
    while True:
        name = _ATTRIBUTE_NAME_RE.search(raw, at)
        if name is None:
            return
        at = name.end()
        opening = _ATTRIBUTE_VALUE_OPENING_RE.match(raw, at)
        if opening is None:
            continue
        start = opening.end()
        end = raw.find('"', start)
        if end < 0:
            return
        at = end + 1
        yield _DocAttribute(
            name.group(0).lower(), js_trim(_decode_entities(raw[start:end])), start, end - start
        )


def _strip_doc_prefix(trimmed: str) -> str:
    rest = trimmed[3:]
    return rest[1:] if rest.startswith(" ") else rest


def line_start_at(source: str, offset: int) -> int:
    """The offset of the start of the line holding `offset`.

    A search from index 0 would find a leading LF and answer 1 for offset 0.
    """
    return 0 if offset <= 0 else source.rfind("\n", 0, offset) + 1


def whole_line_span(source: str, start: int, end: int) -> tuple[int, int]:
    """The span of every physical line [start, end) touches, its newline included."""
    newline = source.find("\n", end)
    return line_start_at(source, start), (len(source) if newline == -1 else newline + 1)


def _is_ordinary_comment(trimmed: str) -> bool:
    return trimmed.startswith("'") and not trimmed.startswith("'''")


def is_xlide_directive_comment(trimmed: str) -> bool:
    """xlide's own directive comments, analysis suppressions and test markers, are
    transparent to the doc scans: a `'''` block attaches to its member through them,
    whatever order the comment grammars stack in. Any OTHER intervening line still
    detaches the block."""
    return _is_ordinary_comment(trimmed) and _XLIDE_DIRECTIVE_RE.match(trimmed) is not None


def attached_comments_start(source: str, decl_start: int) -> int:
    """Where the lines read as part of a declaration start: its `'''` doc comment
    and xlide's directive lines, directly above it in any order. The declaration's
    own line start when it has none.

    Code that moves or deletes the declaration takes these with it; left behind,
    they would belong to whatever declaration came next.
    """
    start = line_start_at(source, decl_start)
    while start > 0:
        previous = line_start_at(source, start - 1)
        trimmed = source[previous : start - 1]
        if trimmed.endswith("\r"):
            trimmed = trimmed[:-1]
        trimmed = trimmed.lstrip()
        if not trimmed.startswith("'''") and not is_xlide_directive_comment(trimmed):
            break
        start = previous
    return start


@dataclass(slots=True)
class DocBlockLine:
    """One `'''` line of a doc block."""

    # Offset of the line's first character.
    start: int
    # Offset of the first of the directive lines directly above it, or `start` when
    # there are none. A line put before this one goes here, so it does not come
    # between a directive and the line the directive is about.
    directives_start: int
    # Offset of the text after `'''` and the one space the parser drops.
    text_start: int
    # That text, to the end of the line.
    text: str


def leading_doc_lines(source: str, decl_start: int) -> list[DocBlockLine]:
    """The `'''` lines directly above a declaration, top to bottom. xlide's own
    directive lines between them are passed over and left out.

    Walks physical lines backward from the declaration's own line; this runs once
    per declaration, so slicing and splitting the whole module prefix here would make
    the pass quadratic in module size.
    """
    line_start = line_start_at(source, decl_start)
    lines: list[DocBlockLine] = []
    while line_start > 0:
        previous_end = line_start - 1  # the '\n' terminating the previous line
        previous_start = line_start_at(source, previous_end)
        line = source[previous_start:previous_end]
        if line.endswith("\r"):
            line = line[:-1]
        trimmed = line.lstrip()
        if is_xlide_directive_comment(trimmed):
            # Suppression and test directives are the product's own grammar; the
            # block attaches through them in any stacking order.
            if lines:
                lines[-1].directives_start = previous_start
            line_start = previous_start
            continue
        if not trimmed.startswith("'''"):
            break
        text = _strip_doc_prefix(trimmed)
        lines.append(
            DocBlockLine(
                start=previous_start,
                directives_start=previous_start,
                text_start=previous_start + len(line) - len(text),
                text=text,
            )
        )
        line_start = previous_start
    lines.reverse()
    return lines


@dataclass(slots=True)
class DocTagOccurrence:
    """One vocabulary tag in a doc block."""

    # The tag in lower case: summary, param, returns, remarks, example or signature.
    tag: str
    # From the `<` to the `>` of the opening tag.
    open: Span
    # True when a `type`, `unit` or `value` attribute says something.
    has_hints: bool = False
    # The `name` attribute, decoded and trimmed, when the tag has a non-empty one.
    name: str | None = None
    # The text between the quotes of that `name` attribute.
    name_span: Span | None = None
    # The tag's text as a hover shows it; None when the tag is never closed.
    text: str | None = None
    # Offset just past the closing tag, or the `/>`; None when never closed.
    end: int | None = None


def scan_doc_tags(lines: list[DocBlockLine]) -> list[DocTagOccurrence] | None:
    """The vocabulary tags of a `'''` block in document order, or None when it has
    none: a block of plain text is a note, which the parser reads as a summary.

    A tag counts as closed the way the parser reads it, by its own closing tag
    before the next tag of the same name opens.
    """
    body = "\n".join(line.text for line in lines)
    if _HAS_TAG_RE.search(body) is None:
        return None
    body_starts: list[int] = []
    at = 0
    for line in lines:
        body_starts.append(at)
        at += len(line.text) + 1

    def to_source(offset: int) -> int:
        # Tags and their closing/name spans need not arrive in offset order: the
        # last line starting at or before this offset, by binary search.
        i = max(0, bisect_right(body_starts, offset) - 1)
        return lines[i].text_start + (offset - body_starts[i])

    lower = body.lower()
    closing_offsets: dict[str, int] = {}
    tags: list[DocTagOccurrence] = []
    at = 0
    while True:
        match = _HAS_TAG_RE.search(body, at)
        if match is None:
            break
        tag = match.group(1).lower()
        angle = body.find(">", match.end())
        if angle < 0:
            break
        open_end = angle + 1
        self_closing = body[angle - 1] == "/"
        at = open_end
        occurrence = DocTagOccurrence(
            tag=tag, open=Span(to_source(match.start()), to_source(open_end))
        )
        # Read the attributes as the parser does: the last of a repeated name wins.
        attrs: dict[str, _DocAttribute] = {}
        attrs_start = match.start() + 1 + len(tag)
        raw_attrs = body[attrs_start : angle - 1 if self_closing else angle]
        for attr in _doc_attributes(raw_attrs):
            attrs[attr.name] = attr
        name = attrs.get("name")
        if name is not None and name.value:
            occurrence.name = name.value
            name_start = attrs_start + name.start
            occurrence.name_span = Span(to_source(name_start), to_source(name_start + name.length))
        occurrence.has_hints = any(
            (hint := attrs.get(key)) is not None and hint.value for key in ("type", "unit", "value")
        )
        if self_closing:
            occurrence.text = ""
            occurrence.end = occurrence.open.end
        else:
            # Opening tags are visited in order. A cached following close stays
            # valid until we pass it; a missing close never needs another scan.
            close = closing_offsets.get(tag)
            if close is None or 0 <= close < open_end:
                close = lower.find(f"</{tag}>", open_end)
                closing_offsets[tag] = close
            if close >= 0:
                reopen = re.compile(rf"<{tag}\b", re.ASCII).search(lower, open_end)
                if reopen is None or close < reopen.start():
                    occurrence.text = _collapse(body[open_end:close])
                    occurrence.end = to_source(close + len(tag) + 3)
        tags.append(occurrence)
    return tags
