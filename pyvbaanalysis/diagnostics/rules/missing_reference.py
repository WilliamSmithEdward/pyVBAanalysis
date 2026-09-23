"""Rule family: naming another application's library without referencing it.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/missingReference.ts.

`Dim doc As Word.Document` in a workbook compiles only when the project references
the Word type library. Without it the VBE refuses the declaration outright, "User-
defined type not defined", and the whole project stops compiling, so this is a real
compile error rather than a style note (oracle case
cross_application_early_binding_without_reference_compile).

EARLY BINDING ONLY. Late binding needs no reference at all: `Dim xl As Object` with
`Set xl = CreateObject("Excel.Application")` names nothing from the library,
resolves through IDispatch at run time, and is the usual way to drive another
application without one. So the rule fires on a name the compiler has to resolve,
a qualified type in an As clause, a New, or a qualified constant, and never on a
string.

It also fires only on the QUALIFIED spelling. An unqualified `Dim wb As Workbook` in
a Word project is indistinguishable from a project class that does not exist yet,
and guessing between the two would put a reference suggestion on ordinary broken
code.

This port adds one gate upstream does not have: the rule speaks only when the
project's reference list is KNOWN. A workbook read from its container carries one;
a loose .bas file does not, and there "the project lacks the Word reference" cannot
be proven, so the rule stays silent rather than guess.
"""

from __future__ import annotations

from ...host.host_libraries import HOST_LIBRARY_NAMES
from ...host.host_model import HostObjectModel
from ...lexer.token_helpers import token_name
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn
from ..model import VbaAddLibraryReferenceData, VbaDiagnosticData

# The libraries a project can be given a reference to: lowercased name -> as written.
_ADDABLE = {
    HOST_LIBRARY_NAMES[token].lower(): HOST_LIBRARY_NAMES[token]
    for token in ("excel", "word", "powerpoint", "access")
}


def _libraries_in_model(model: HostObjectModel) -> set[str]:
    """Which libraries the model can already answer for, from its own type keys."""
    out: set[str] = set()
    for qualified in model.get("types") or {}:
        library, dot, _ = qualified.partition(".")
        if dot and library:
            out.add(library.lower())
    return out


def _qualified_names_in(source: str) -> list[tuple[str, Span]]:
    """Every `Library.Member` in the module where the compiler has to resolve
    `Library`: in an As clause, after New, or standing as a value.

    The whole module is scanned rather than each procedure's statements, because
    `Dim xl As Excel.Application`, the commonest early binding there is, and a
    module-level `Private mApp As Excel.Application` with it, is a declaration, and
    the per-statement walk does not reach declarations.

    A string literal is one token, so `CreateObject("Excel.Application")` is not a
    match: late binding names nothing the compiler has to resolve.
    """
    # The module's shared tokenization, not a second lex of the whole module.
    toks = [
        t
        for t in tokenize_cached(source)
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    out: list[tuple[str, Span]] = []
    for i in range(len(toks) - 2):
        library = token_name(toks[i])
        if toks[i].kind is not TokenKind.IDENTIFIER or not library:
            continue
        if toks[i + 1].raw_text != ".":
            continue
        if toks[i + 2].kind not in (TokenKind.IDENTIFIER, TokenKind.KEYWORD):
            continue
        # A member access further along a chain (`a.b.c`) is not a library
        # qualifier: `b` there is a member of whatever `a` is.
        if i > 0 and toks[i - 1].raw_text == ".":
            continue
        out.append((library, Span(toks[i].start, toks[i + 2].end)))
    return out


def libraries_named_in(source: str) -> set[str]:
    """The libraries the module names early bound, lowercased.

    Removing a reference is the other half of this rule: a project can be told which
    of its modules would stop compiling before the reference goes, which is what the
    VBE's own Tools > References dialog never says.
    """
    return {library.lower() for library, _ in _qualified_names_in(source)}


def check_missing_library_reference(
    source: str,
    model: HostObjectModel,
    references_known: bool,
    push: PushFn,
) -> None:
    """Module rule: a type or constant qualified with an Office library the project
    does not reference.

    Reported once per library per module. A project missing a reference names it on
    every line that uses it, and one diagnostic per line would bury the module in the
    same message with the same one fix.
    """
    if not references_known:
        return
    present = _libraries_in_model(model)
    # Nothing is known about any library, so nothing can be said about one being
    # absent. A model with no qualified types gets silence.
    if not present:
        return
    seen: set[str] = set()
    for found_library, span in _qualified_names_in(source):
        lower = found_library.lower()
        library = _ADDABLE.get(lower)
        if library is None or lower in present or lower in seen:
            continue
        seen.add(lower)
        push(
            "missingLibraryReference",
            f"'{library}' is not referenced by this project, so {library}.* cannot be "
            f"resolved. Add a reference to the {library} object library, or use late "
            f'binding: Dim x As Object: Set x = CreateObject("{library}.Application").',
            span,
            VbaDiagnosticData(add_library_reference=VbaAddLibraryReferenceData(library=lower)),
        )
