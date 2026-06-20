"""Inline diagnostic suppression via comment directives.

A ``'@pyvba-ignore`` directive in the VBA source suppresses diagnostics. There are
three forms (each is a single-apostrophe comment; ``'''`` doc comments and ``Rem``
comments are never directives):

    x = Foo()                 '@pyvba-ignore: unknown-call
    '@pyvba-ignore-next-line: undeclared-variable
    n = Bar
    '@pyvba-ignore-file: option-explicit-missing   (before the first source line)

* ``'@pyvba-ignore`` suppresses diagnostics on the directive's own physical line.
* ``'@pyvba-ignore-next-line`` suppresses the next physical line.
* ``'@pyvba-ignore-file`` suppresses the whole module; it must appear before the
  first non-comment, non-attribute source line.

The code list is optional (omitted, or the word ``all``, means every code),
comma-separated and matched case-insensitively; a ``-- reason`` trailer is free text.
A malformed directive (an unknown verb under the ``@pyvba-ignore-`` namespace, an
unknown code, or a misplaced ``-file`` directive) is reported as
``analysis-suppression-directive`` and suppresses nothing.

Suppression is lexical and conditional-compilation-agnostic, applied as a
post-analysis filter. It covers the rule-catalogue codes; the two structural codes
(missing-block-closer, unmatched-block-closer) are emitted by the parser, not this
engine, so they are not in its output and cannot be suppressed here. Member and block
scopes are intentionally out of scope. A directive never suppresses the
analysis-suppression-directive diagnostic itself.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..lexer.token_kinds import TokenKind
from ..lexer.tokenize import tokenize
from ..parser.nodes import Span
from .model import VbaDiagnostic, line_col
from .rule_metadata import diagnostic_metadata_for_code

DIRECTIVE_DIAGNOSTIC_CODE = "analysis-suppression-directive"

# A comment whose body (after the apostrophe) opens the @pyvba-ignore namespace; the
# boundary lookahead keeps an ordinary comment like "@pyvba-ignored idea" from matching.
_CANDIDATE_RE = re.compile(r"^@pyvba-ignore(?=$|[-\s:])", re.IGNORECASE)
# A well-formed directive: the namespace, an optional verb, then end / space / colon.
_DIRECTIVE_RE = re.compile(
    r"^@pyvba-ignore(-next-line|-file)?(?=$|[\s:])(.*)$", re.IGNORECASE | re.DOTALL
)


@dataclass(slots=True)
class _Target:
    """The codes a directive suppresses for one scope (a line, or the whole file)."""

    all_codes: bool = False
    codes: set[str] = field(default_factory=set)

    def add(self, codes: set[str] | None) -> None:
        if codes is None:
            self.all_codes = True
        else:
            self.codes |= codes

    def covers(self, code: str) -> bool:
        return self.all_codes or code in self.codes


@dataclass(slots=True)
class InlineSuppressionScan:
    """The suppression state parsed from a module's directive comments."""

    file_target: _Target | None = None
    line_targets: dict[int, _Target] = field(default_factory=dict)
    # (span, message) for each malformed directive, to surface as a diagnostic.
    issues: list[tuple[Span, str]] = field(default_factory=list)


def scan_inline_suppressions(source: str) -> InlineSuppressionScan:
    """Parse the ``'@pyvba-ignore`` directives in ``source`` into suppression state."""
    scan = InlineSuppressionScan()
    first_source_line = _first_source_line(source)
    for token in tokenize(source):
        if token.kind is not TokenKind.COMMENT:
            continue
        raw = token.raw_text
        if not raw.startswith("'") or raw.startswith("'''"):
            continue  # Rem comments and doc comments are not directives
        body = raw[1:].lstrip()
        if _CANDIDATE_RE.match(body) is None:
            continue
        span = Span(token.start, token.end)
        directive = _DIRECTIVE_RE.match(body)
        if directive is None:
            scan.issues.append((span, f"Unknown suppression directive: {body.split()[0]!r}."))
            continue
        verb = directive.group(1)
        codes = _parse_codes(directive.group(2), span, scan)
        line = line_col(source, token.start)[0]
        if verb is None:
            scan.line_targets.setdefault(line, _Target()).add(codes)
        elif verb.lower() == "-next-line":
            scan.line_targets.setdefault(line + 1, _Target()).add(codes)
        else:  # -file
            if first_source_line is not None and line >= first_source_line:
                scan.issues.append(
                    (span, "'@pyvba-ignore-file must appear before the first source line.")
                )
                continue
            if scan.file_target is None:
                scan.file_target = _Target()
            scan.file_target.add(codes)
    return scan


def filter_inline_suppressions(
    source: str, diagnostics: list[VbaDiagnostic], scan: InlineSuppressionScan
) -> list[VbaDiagnostic]:
    """Drop the diagnostics a scan suppresses, never the directive diagnostic itself."""
    if scan.file_target is None and not scan.line_targets:
        return diagnostics
    kept: list[VbaDiagnostic] = []
    for diag in diagnostics:
        if diag.code == DIRECTIVE_DIAGNOSTIC_CODE:
            kept.append(diag)
            continue
        if scan.file_target is not None and scan.file_target.covers(diag.code):
            continue
        target = scan.line_targets.get(line_col(source, diag.span.start)[0])
        if target is not None and target.covers(diag.code):
            continue
        kept.append(diag)
    return kept


def _parse_codes(rest: str, span: Span, scan: InlineSuppressionScan) -> set[str] | None:
    """Parse an optional ``: code, code -- reason`` tail; None means every code."""
    rest = rest.strip()
    if rest.startswith(":"):
        rest = rest[1:].strip()
    if "--" in rest:
        rest = rest.split("--", 1)[0].strip()
    if rest == "" or rest.lower() == "all":
        return None
    codes: set[str] = set()
    for part in rest.split(","):
        code = part.strip().lower()
        if code == "":
            scan.issues.append((span, "Empty code in suppression directive."))
        elif code == "all":
            scan.issues.append((span, "'all' cannot be combined with specific codes."))
        elif diagnostic_metadata_for_code(code) is None:
            scan.issues.append(
                (span, f"Unknown diagnostic code {code!r} in suppression directive.")
            )
        else:
            codes.add(code)
    return codes


def _first_source_line(source: str) -> int | None:
    """1-based line of the first non-comment, non-attribute source line, or None."""
    for i, line in enumerate(source.splitlines(), start=1):
        s = line.strip()
        if not s or s.startswith("'"):
            continue
        if len(s) >= 3 and s[:3].lower() == "rem" and (len(s) == 3 or not (s[3].isalnum() or s[3] == "_")):
            continue
        if s.lower().startswith("attribute "):
            continue
        return i
    return None
