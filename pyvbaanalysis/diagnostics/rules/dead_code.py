"""Dead code: declarations nothing reads, private procedures nothing calls, and
statements nothing can reach.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/deadCode.ts.

Every finding here is a structural fact about the module's own text, not a guess
about how it runs, and each one stops at the point where the text can no longer
prove it:

- A local or module-private variable is "unused" when no token in its scope names
  it, and "never read" when every token that does is a plain assignment to it.
  Passing it to a procedure counts as a read (ByRef may fill it), so does indexing
  it, a member access on it, a `For` loop over it, and any mention inside an
  inactive `#If` arm.
- Only Private procedures are reported as uncalled. A Public one may be wired to a
  button, a shape, the ribbon, a hotkey, `OnTime` or `Application.Run` in a file
  the analyzer cannot see, so its silence proves nothing. A Private one can only be
  reached from its own module, or by name in a string, which is why every string
  literal in the project is searched too. Event handlers and procedures with member
  attributes are never reported: the host calls those.
- Code after `Exit Sub`, `Exit Function`, `Exit Property`, `Exit Do`, `Exit For`,
  `GoTo`, `Resume`, `Return` or `End` in the same block is unreachable until a
  label, a line number, a `Case` or a `#If` gives the flow somewhere to land.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Literal

from ...completion.event_handlers import event_handler_procedure_for_name
from ...conditional import ConditionalActivityTracker
from ...docs.doc_comment import attached_comments_start, whole_line_span
from ...lexer.token_helpers import identifier_words
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    IfBlockNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
)
from ...references import ReferenceKind, classify_reference_kinds
from ...symbols.symbol_model import ModuleSymbolKind, ModuleSymbols, VbaSymbol, is_procedure_kind
from ..context import PushFn, statement_tokens
from ..model import (
    VbaDiagnosticData,
    VbaEdit,
    VbaRemoveDeclarationData,
    VbaRemoveUnreachableCodeData,
)
from ..walker import (
    active_module_members,
    declared_name_span,
    for_each_variable_group,
    is_inactive_node,
    token_text,
)

# ------------------------------------------------------------ unused names


@dataclass(slots=True)
class _TrackedDeclaration:
    name: str
    lower: str
    name_span: Span
    group: VariableGroupNode
    scope: Literal["local", "module"]


@dataclass(slots=True)
class _Reference:
    offset: int
    # Read by construction: a `For`/`For Each` control variable is read by the loop
    # itself, and `x(i) = v` reads the array it stores into.
    read_by_form: bool


def check_unused_declarations(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Locals nothing uses, module-private variables and constants nothing uses, and
    variables whose only mentions assign to them."""
    tokens = tokenize_cached(source)
    attributed = {symbol.name_span.start for symbol in symbols.all if symbol.attributes}

    members = active_module_members(mod, activity)
    procedures = [member for member in members if isinstance(member, ProcedureNode)]

    # Module-level declarations that only this module can name.
    module_level: list[_TrackedDeclaration] = []
    for member in members:
        if not isinstance(member, VariableGroupNode) or member.with_events or not _is_module_private(member):
            continue
        for decl in member.declarations:
            if decl.name_span is None or decl.name_span.start in attributed:
                continue
            module_level.append(
                _TrackedDeclaration(decl.name, decl.name.lower(), decl.name_span, member, "module")
            )

    # Procedures whose parameters or locals shadow a module-level name: a mention
    # inside them binds to the shadow, not the module variable.
    shadowing: dict[str, list[Span]] = {}
    for proc in procedures:
        names = {param.name.lower() for param in proc.params}

        def collect(group: VariableGroupNode, _names: set[str] = names) -> None:
            _names.update(decl.name.lower() for decl in group.declarations)

        for_each_variable_group(proc.body, collect, activity)
        for lower in names:
            shadowing.setdefault(lower, []).append(proc.span)

    # One pass over the token stream serves every scope: the module-level names see
    # the whole module, each procedure's locals see its own span. Reference kinds
    # are then classified in one walk as well, so the cost stays linear in module
    # size however many procedures it has.
    scopes: list[tuple[int, int, list[_TrackedDeclaration]]] = []
    if module_level:
        scopes.append((0, len(source), module_level))
    for proc in procedures:
        locals_: list[_TrackedDeclaration] = []

        def track(group: VariableGroupNode, _locals: list[_TrackedDeclaration] = locals_) -> None:
            for decl in group.declarations:
                if decl.name_span is None or decl.name_span.start in attributed:
                    continue
                _locals.append(_TrackedDeclaration(decl.name, decl.name.lower(), decl.name_span, group, "local"))

        for_each_variable_group(proc.body, track, activity)
        if locals_:
            scopes.append((proc.span.start, proc.span.end, locals_))
    if not scopes:
        return

    references_by_scope = [
        _collect_references(
            tokens,
            start,
            end,
            {decl.name_span.start for decl in declarations},
            {decl.lower for decl in declarations},
        )
        for start, end, declarations in scopes
    ]
    all_offsets = [
        ref.offset for references in references_by_scope for refs in references.values() for ref in refs
    ]
    kinds = classify_reference_kinds(source, all_offsets)

    for (_start, _end, declarations), references in zip(scopes, references_by_scope):
        for tracked in declarations:
            own = references.get(tracked.lower, [])
            if tracked.scope == "module":
                shadows = shadowing.get(tracked.lower, [])
                own = [
                    ref for ref in own
                    if not any(span.start <= ref.offset < span.end for span in shadows)
                ]
            _report(source, tracked, own, kinds, push)


def _is_module_private(group: VariableGroupNode) -> bool:
    modifier = group.modifier.lower()
    if modifier in ("private", "dim"):
        return True
    # A bare `Const` at module level is private by default (MS-VBAL 5.2.3.2).
    return group.is_const and modifier == ""


def _report(
    source: str,
    decl: _TrackedDeclaration,
    references: list[_Reference],
    kinds: Mapping[int, ReferenceKind],
    push: PushFn,
) -> None:
    if not references:
        if decl.group.is_const:
            what = "Constant"
        elif decl.scope == "module":
            what = "Module-level variable"
        else:
            what = "Local variable"
        removal = _remove_declaration_data(source, decl)
        push(
            "unusedVariable",
            f"{what} '{decl.name}' is declared but never used.",
            decl.name_span,
            VbaDiagnosticData(remove_declaration=removal) if removal is not None else None,
        )
        return
    if decl.group.is_const:
        return
    ever_written = any(kinds.get(ref.offset) == "write" and not ref.read_by_form for ref in references)
    ever_read = any(ref.read_by_form or kinds.get(ref.offset) != "write" for ref in references)
    if ever_written and not ever_read:
        push(
            "variableNeverRead",
            f"Variable '{decl.name}' is assigned but its value is never read.",
            decl.name_span,
        )


def _collect_references(
    tokens: Sequence[VbaToken],
    start: int,
    end: int,
    declared: AbstractSet[int],
    names: AbstractSet[str],
) -> dict[str, list[_Reference]]:
    """Every mention of one of `names` within [start, end) that could bind to a
    variable: not a member name, not a named-argument name, not a declaration site.
    Keyed by lowercased name.

    Upstream collects every name in the scope. Only the scope's own declarations
    are ever looked up, so this collects just those: every other mention would also
    have been classified as a read or a write, and on a large module that work
    cost more than the rest of the rule."""
    out: dict[str, list[_Reference]] = {}
    first = _first_token_at_or_after(tokens, start)
    prev = tokens[first - 1] if first >= 1 else None
    prev2 = tokens[first - 2] if first >= 2 else None
    for i in range(first, len(tokens)):
        token = tokens[i]
        if token.start >= end:
            break
        if token.kind is TokenKind.IDENTIFIER or token.kind is TokenKind.KEYWORD:
            word = token_text(token)
            following = tokens[i + 1] if i + 1 < len(tokens) else None
            if (
                word in names
                and token.start not in declared
                and not _is_member_name(prev)
                and not _is_named_argument(following)
            ):
                prev_word = token_text(prev)
                read_by_form = (
                    prev_word == "for"
                    or (prev_word == "each" and token_text(prev2) == "for")
                    or (
                        following is not None
                        and following.kind is TokenKind.PUNCTUATION
                        and following.raw_text == "("
                    )
                )
                out.setdefault(word, []).append(_Reference(token.start, read_by_form))
        if token.kind is not TokenKind.COMMENT:
            prev2 = prev
            prev = token
    return out


def _first_token_at_or_after(tokens: Sequence[VbaToken], offset: int) -> int:
    """Index of the first token starting at or after `offset` (tokens are in source
    order)."""
    low, high = 0, len(tokens)
    while low < high:
        mid = (low + high) >> 1
        if tokens[mid].start < offset:
            low = mid + 1
        else:
            high = mid
    return low


def _is_member_name(prev: VbaToken | None) -> bool:
    return prev is not None and (
        (prev.kind is TokenKind.PUNCTUATION and prev.raw_text == ".")
        or (prev.kind is TokenKind.OPERATOR and prev.raw_text == "!")
    )


def _is_named_argument(following: VbaToken | None) -> bool:
    return following is not None and following.kind is TokenKind.OPERATOR and following.raw_text == ":="


def _remove_declaration_data(source: str, decl: _TrackedDeclaration) -> VbaRemoveDeclarationData | None:
    """The edit that removes one declaration: its whole line when it stands alone
    there, else its own name (and separator) from a `Dim a, b` list. None when the
    statement shares its line with something else."""
    group = decl.group
    line_start, line_end = whole_line_span(source, group.span.start, group.span.end)
    line_text = source[line_start:line_end]
    if len(group.declarations) == 1:
        # Only the declaration (and perhaps a comment) on its line.
        statement_text = source[group.span.start : group.span.end]
        rest = re.sub(r"\r?\n$", "", line_text.replace(statement_text, "", 1)).strip()
        if rest != "" and not rest.startswith("'"):
            return None
        if "\n" in statement_text:
            return None
        # A module variable's doc comment goes with it; left behind, it would
        # document the declaration that came next.
        start = attached_comments_start(source, group.span.start) if decl.scope == "module" else line_start
        return VbaRemoveDeclarationData(decl.name, VbaEdit(Span(start, line_end), ""))
    index = next(
        (
            i
            for i, candidate in enumerate(group.declarations)
            if candidate.name_span is not None and candidate.name_span.start == decl.name_span.start
        ),
        -1,
    )
    if index < 0:
        return None
    own = group.declarations[index]
    if index < len(group.declarations) - 1:
        following = group.declarations[index + 1]
        return VbaRemoveDeclarationData(decl.name, VbaEdit(Span(own.span.start, following.span.start), ""))
    previous = group.declarations[index - 1]
    return VbaRemoveDeclarationData(decl.name, VbaEdit(Span(previous.span.end, own.span.end), ""))


# ---------------------------------------------------- uncalled procedures

_AUTO_MACRO = re.compile(r"^auto_(open|close|activate|deactivate|exec|new|add|remove)$", re.IGNORECASE)
_ATTRIBUTE_LINE = re.compile(r"^\s*Attribute\b", re.IGNORECASE)


def check_unused_private_procedures(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    module_kind: ModuleSymbolKind,
    activity: ConditionalActivityTracker | None,
    project_string_literal_words: AbstractSet[str] | None,
    push: PushFn,
) -> None:
    """Private procedures no token in the module and no string in the project names."""
    # The procedure symbol at each start offset, first one wins as upstream's
    # `symbols.all.find` does. Built once: looking it up per candidate by scanning the
    # symbol list is quadratic, and a large class has a thousand private procedures.
    procedure_symbol_at: dict[int, VbaSymbol] = {}
    for symbol in symbols.all:
        if is_procedure_kind(symbol.kind):
            procedure_symbol_at.setdefault(symbol.full_span.start, symbol)
    candidates = [
        member
        for member in active_module_members(mod, activity)
        if isinstance(member, ProcedureNode)
        and any(modifier.lower() == "private" for modifier in member.modifiers)
        and member.name_span is not None
        and not _has_member_attribute(source, member, procedure_symbol_at)
        and not _is_host_called(member.name, module_kind)
    ]
    if not candidates:
        return
    declaration_sites = {
        member.name_span.start
        for member in mod.members
        if isinstance(member, ProcedureNode) and member.name_span is not None
    }
    # Every mention of a name, by offset: a Function assigning its own return value
    # names itself, and a procedure calling only itself is still dead, so mentions
    # inside the procedure's own body do not count for it.
    mentions: dict[str, list[int]] = {}
    own_string_words: set[str] = set()
    for token in tokenize_cached(source):
        if token.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD) and token.start not in declaration_sites:
            mentions.setdefault(token_text(token), []).append(token.start)
        elif token.kind is TokenKind.STRING_LITERAL and project_string_literal_words is None:
            own_string_words.update(identifier_words(token.raw_text))
    string_words = project_string_literal_words if project_string_literal_words is not None else own_string_words

    for proc in candidates:
        lower = proc.name.lower()
        outside = any(
            offset < proc.span.start or offset >= proc.span.end for offset in mentions.get(lower, ())
        )
        if outside or lower in string_words:
            continue
        if proc.proc_kind is ProcKind.SUB:
            message = f"Private Sub '{proc.name}' is never called."
        elif proc.proc_kind is ProcKind.FUNCTION:
            message = f"Private Function '{proc.name}' is never called."
        else:
            message = f"Private Property '{proc.name}' is never used."
        push("unusedProcedure", message, declared_name_span(source, proc.span, proc.name))


def _has_member_attribute(
    source: str, proc: ProcedureNode, procedure_symbol_at: Mapping[int, VbaSymbol]
) -> bool:
    """Whether a procedure carries an `Attribute` line: a hotkey, a description, a
    default member. An exported module writes the line inside the procedure under
    its header; the symbol table attaches one written after it."""
    if proc.attributes:
        return True
    symbol = procedure_symbol_at.get(proc.span.start)
    if symbol is not None and symbol.attributes:
        return True
    return any(
        isinstance(node, StatementNode) and _ATTRIBUTE_LINE.match(source[node.span.start : node.span.end])
        for node in proc.body
    )


def _is_host_called(name: str, module_kind: ModuleSymbolKind) -> bool:
    """Whether the host, not code, calls a procedure by this name: an event handler
    (`Worksheet_Change`, `CommandButton1_Click`, `Class_Initialize`, an interface
    member `IFoo_Bar`) in an object module, or an `Auto_Open` style macro in a
    standard one."""
    if event_handler_procedure_for_name(name) is not None:
        return True
    if module_kind is ModuleSymbolKind.STANDARD:
        return _AUTO_MACRO.match(name) is not None
    return "_" in name


# ------------------------------------------------------- unreachable code


def check_unreachable_code(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Statements after an unconditional exit in the same block, until a landing
    point."""

    def walk_body(body: Sequence[BodyNode]) -> None:
        terminator: str | None = None
        dead: Span | None = None

        def flush() -> None:
            nonlocal terminator, dead
            if dead is not None and terminator is not None:
                start, end = whole_line_span(source, dead.start, dead.end)
                push(
                    "unreachableCode",
                    f"Unreachable code after '{terminator}'.",
                    dead,
                    VbaDiagnosticData(
                        remove_unreachable_code=VbaRemoveUnreachableCodeData(VbaEdit(Span(start, end), ""))
                    ),
                )
            dead = None
            terminator = None

        for node in body:
            if is_inactive_node(activity, node):
                continue
            if isinstance(node, ConditionalDirectiveNode):
                flush()
                continue
            if is_leaf_statement(node) and node.single_line_if_tail:
                # It runs only with its single-line If's branch (MS-VBAL 5.4.2.9):
                # an Exit there ends nothing, and it is dead when its If is.
                if terminator is not None:
                    dead = Span(dead.start if dead is not None else node.span.start, node.span.end)
                continue
            if is_leaf_statement(node):
                toks = statement_tokens(source, node.span)
                if _is_landing_point(source, node, toks):
                    flush()
                    exit_text = _terminal_statement(_tokens_after_line_number(toks))
                    if exit_text is not None:
                        terminator = exit_text
                    continue
                if terminator is not None:
                    dead = Span(dead.start if dead is not None else node.span.start, node.span.end)
                    continue
                exit_text = _terminal_statement(toks)
                if exit_text is not None:
                    terminator = exit_text
                continue
            # A block node.
            if terminator is not None:
                if _block_has_landing_point(source, node):
                    flush()
                    walk_block(node)
                    continue
                dead = Span(dead.start if dead is not None else node.span.start, node.span.end)
                continue
            walk_block(node)
        flush()

    def walk_block(node: BodyNode) -> None:
        if isinstance(node, IfBlockNode):
            for branch in node.branches:
                walk_body(branch.body)
            return
        child = getattr(node, "body", None)
        if isinstance(child, list):
            walk_body(child)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            walk_body(member.body)


def _is_label_line(source: str, node: BodyNode, toks: Sequence[VbaToken]) -> bool:
    return (
        isinstance(node, StatementNode)
        and len(toks) == 1
        and toks[0].kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD)
        and source[node.span.end : node.span.end + 1] == ":"
    )


def _is_landing_point(source: str, node: BodyNode, toks: Sequence[VbaToken]) -> bool:
    """A label, a line number, or a `Case` arm: somewhere control can arrive."""
    if not toks:
        return False
    if toks[0].kind is TokenKind.INTEGER_LITERAL:
        return True
    if _is_label_line(source, node, toks):
        return True
    return token_text(toks[0]) in ("case", "else", "elseif")


def _tokens_after_line_number(toks: Sequence[VbaToken]) -> Sequence[VbaToken]:
    return toks[1:] if toks and toks[0].kind is TokenKind.INTEGER_LITERAL else toks


def _terminal_statement(toks: Sequence[VbaToken]) -> str | None:
    if not toks:
        return None
    head = token_text(toks[0])
    second = token_text(toks[1]) if len(toks) > 1 else ""
    if head == "exit" and second in ("sub", "function", "property", "do", "for"):
        return f"Exit {toks[1].canonical_text or toks[1].raw_text}"
    if head == "goto" and len(toks) >= 2:
        return f"GoTo {toks[1].raw_text}"
    if head == "resume":
        return "Resume" if len(toks) == 1 else f"Resume {toks[1].canonical_text or toks[1].raw_text}"
    if head == "end" and len(toks) == 1:
        return "End"
    if head == "return" and len(toks) == 1:
        return "Return"
    return None


def _block_has_landing_point(source: str, node: BodyNode) -> bool:
    bodies: list[Sequence[BodyNode]] = []
    if isinstance(node, IfBlockNode):
        bodies.extend(branch.body for branch in node.branches)
    else:
        child = getattr(node, "body", None)
        if isinstance(child, list):
            bodies.append(child)
    for body in bodies:
        for child_node in body:
            if is_leaf_statement(child_node):
                toks = statement_tokens(source, child_node.span)
                if toks and toks[0].kind is TokenKind.INTEGER_LITERAL:
                    return True
                if _is_label_line(source, child_node, toks):
                    return True
            elif isinstance(child_node, ConditionalDirectiveNode):
                return True
            elif _block_has_landing_point(source, child_node):
                return True
    return False
