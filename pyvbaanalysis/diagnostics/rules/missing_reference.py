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

from collections.abc import Sequence
from collections.abc import Set as AbstractSet

from ...host.host_libraries import HOST_LIBRARY_NAMES
from ...host.host_model import HostObjectModel
from ...lexer.token_helpers import token_name
from ...lexer.token_kinds import TokenKind
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import Span
from ..context import PushFn
from ..model import VbaAddLibraryReferenceData, VbaDiagnosticData
from ...symbols.build_module_symbols import build_module_symbols
from ...symbols.symbol_model import ModuleSymbols, ModuleSymbolKind, VbaSymbol, is_procedure_kind
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionInput, BareIdentifierResolutionScope, resolve_bare_identifier_binding
from ...identity_cache import IdentityLru

_MODEL_LIBRARIES = IdentityLru()

# The libraries a project can be given a reference to: lowercased name -> as written.
_ADDABLE = {
    HOST_LIBRARY_NAMES[token].lower(): HOST_LIBRARY_NAMES[token]
    for token in ("excel", "word", "powerpoint", "access")
}


def _libraries_in_model(model: HostObjectModel) -> set[str]:
    """Which libraries the model can already answer for, from its own type keys."""
    cached = _MODEL_LIBRARIES.get(model)
    if cached is not None:
        return cached  # type: ignore[no-any-return]
    out: set[str] = set()
    for qualified in model.get("types") or {}:
        library, dot, _ = qualified.partition(".")
        if dot and library:
            out.add(library.lower())
    return _MODEL_LIBRARIES.put(out, model)  # type: ignore[no-any-return]


def _qualified_names_in(source: str, symbols: ModuleSymbols | None = None, project_visible_symbols: Sequence[VbaSymbol] | None = None) -> list[tuple[str, Span, bool]]:
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
    tokens = tokenize_cached(source)
    symbols = symbols or build_module_symbols("", ModuleSymbolKind.STANDARD, source)
    procedures = [symbol for symbol in symbols.root.children or [] if is_procedure_kind(symbol.kind)]
    implicit_writes: dict[int, set[str]] = {}
    procedure_index = 0
    previous = ""
    for i, tok in enumerate(tokens):
        if tok.kind is TokenKind.COMMENT:
            continue
        while procedure_index < len(procedures) and procedures[procedure_index].full_span.end < tok.start:
            procedure_index += 1
        procedure = procedures[procedure_index] if procedure_index < len(procedures) else None
        if procedure and procedure.full_span.start <= tok.start and tok.kind is TokenKind.IDENTIFIER and i + 1 < len(tokens) and tokens[i + 1].raw_text == "=" and previous in ("", "set", "let", "then", "else"):
            implicit_writes.setdefault(procedure.full_span.start, set()).add(tok.raw_text.lower())
        previous = "" if tok.kind is TokenKind.NEWLINE or tok.raw_text == ":" else tok.raw_text.lower()
    procedure_index = 0
    toks = [
        t
        for t in tokens
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    out: list[tuple[str, Span, bool]] = []
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
        type_qualifier = i > 0 and toks[i - 1].raw_text.lower() in ("as", "new", "implements")
        if not type_qualifier:
            while procedure_index < len(procedures) and procedures[procedure_index].full_span.end < toks[i].start:
                procedure_index += 1
            procedure = procedures[procedure_index] if procedure_index < len(procedures) else None
            binding = resolve_bare_identifier_binding(BareIdentifierResolutionInput(current_module=symbols, project_visible_symbols=project_visible_symbols or (), enclosing_procedure=procedure if procedure and procedure.full_span.start <= toks[i].start else None, name=library, context=BareIdentifierContext.MEMBER_RECEIVER, offset=toks[i].start))
            if binding.scope is not BareIdentifierResolutionScope.UNRESOLVED or (procedure and library.lower() in implicit_writes.get(procedure.full_span.start, set())):
                continue
        out.append((library, Span(toks[i].start, toks[i + 2].end), type_qualifier))
    return out


def libraries_named_in(source: str) -> set[str]:
    """The libraries the module names early bound, lowercased.

    Removing a reference is the other half of this rule: a project can be told which
    of its modules would stop compiling before the reference goes, which is what the
    VBE's own Tools > References dialog never says.
    """
    return {library.lower() for library, _, _ in _qualified_names_in(source)}


# The Scripting library's types a module names unqualified, which no default
# reference brings.
_SCRIPTING_TYPES: frozenset[str] = frozenset({"dictionary", "filesystemobject", "textstream"})


def check_missing_scripting_reference(
    source: str,
    referenced_libraries: Sequence[str] | None,
    project_types: AbstractSet[str],
    push: PushFn,
) -> None:
    """Module rule: an early-bound Scripting type in a project whose references are
    known and do not include the Scripting Runtime.

    `Dim d As Scripting.Dictionary` and `Dim d As New Dictionary` then do not
    compile, "User-defined type not defined" (issue #349, measured in Excel 16.0).
    Silent when the references are not known, and for a name the project declares
    itself. Reported once per module, as a missing library is.
    """
    if referenced_libraries is None or any(name.lower() == "scripting" for name in referenced_libraries):
        return
    toks = [
        t
        for t in tokenize_cached(source)
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]

    def raw_at(index: int) -> str | None:
        return toks[index].raw_text if index < len(toks) else None

    i = 0
    while i + 1 < len(toks):
        word = toks[i].raw_text.lower()
        if word != "as" and word != "new":
            i += 1
            continue
        at = i + 1
        if word == "as" and (raw_at(at) or "").lower() == "new":
            at += 1
        qualified = (
            (raw_at(at) or "").lower() == "scripting"
            and raw_at(at + 1) == "."
            and token_name(toks[at + 2] if at + 2 < len(toks) else None) is not None
        )
        name = (
            token_name(toks[at + 2])
            if qualified
            else token_name(toks[at] if at < len(toks) else None)
        )
        lower = name.lower() if name is not None else None
        if (
            not name
            or not lower
            or (
                not qualified
                and (lower not in _SCRIPTING_TYPES or lower in project_types or raw_at(at + 1) == ".")
            )
        ):
            i += 1
            continue
        push(
            "missingLibraryReference",
            f"'{name}' is the Scripting Runtime's, which this project does not reference. Add a "
            "reference to Microsoft Scripting Runtime, or bind late: "
            f'Dim x As Object: Set x = CreateObject("Scripting.{name}"). This is a VBE compile '
            "error: User-defined type not defined.",
            Span(toks[at].start, (toks[at + 2] if qualified else toks[at]).end),
        )
        return


def check_missing_library_reference(
    source: str,
    model: HostObjectModel,
    references_known: bool,
    push: PushFn,
    project_modules: AbstractSet[str] = frozenset(),
    symbols: ModuleSymbols | None = None,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
) -> None:
    """Module rule: a type or constant qualified with an Office library the project
    does not reference.

    Reported once per library per module. A project missing a reference names it on
    every line that uses it, and one diagnostic per line would bury the module in the
    same message with the same one fix.

    `project_modules` is the project's module names, lowercased: a module named Word
    is called as Word.Hi (issue #357).
    """
    if not references_known:
        return
    present = _libraries_in_model(model)
    # Nothing is known about any library, so nothing can be said about one being
    # absent. A model with no qualified types gets silence.
    if not present:
        return
    seen: set[str] = set()
    explicit = any(tok.kind is TokenKind.KEYWORD and tok.raw_text.lower() == "explicit" for tok in tokenize_cached(source))
    for found_library, span, type_qualifier in _qualified_names_in(source, symbols, project_visible_symbols):
        if not type_qualifier and not explicit:
            continue
        lower = found_library.lower()
        library = _ADDABLE.get(lower)
        if library is None or lower in present or lower in project_modules or lower in seen:
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
