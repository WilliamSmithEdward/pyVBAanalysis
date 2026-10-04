"""Helpers shared by more than one diagnostics rule family.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/shared.ts: the exhaustive
member-surface resolver, the read-reference scanner, the repeated-key reporter,
and the statement classifiers several rules use.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import Protocol, TypeVar

from ...call.call_context import bare_call_statement_target as call_statement_target
from ...completion import MemberCompletionContext
from ...completion.member_access import (
    ExhaustiveMemberSurface,
    project_class_members_index,
    resolve_exhaustive_member_surface_at,
)
from ...conditional import ConditionalActivityTracker, collect_conditional_directives, inactive_node_skip
from ...constants.integer_constant_expression import bankers_round
from ...js_compat import js_trim
from ...lexer.keyword_table import is_reserved_identifier
from ...lexer.token_helpers import is_decimal_line_number
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    DoBlockNode,
    ForBlockNode,
    ModuleNode,
    Span,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.symbol_model import VbaProjectClassMembers, qualified_procedure_key
from ..call_extraction import CallableTypeSignature
from ..context import statement_tokens
from ..walker import (
    absolute_span,
    block_footer_line_span,
    block_header_line_span,
    first_executable_token_index,
    raw_expression_tokens,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
    top_level_operator_index,
)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


# -- exhaustive member-surface resolver ------------------------------------


def resolve_exhaustive_member_surface(
    source: str, dot_end_offset: int, member_ctx: MemberCompletionContext
) -> ExhaustiveMemberSurface | None:
    """The member surface at a member-access dot, but ONLY when it is exhaustive.

    Ported from resolveExhaustiveMemberSurface (shared.ts). Returns None unless the
    resolved receiver type yields an exhaustive surface, the no-false-positive gate
    for member-not-found: a non-exhaustive host type, Object/Variant, or unresolved
    receiver produces no surface, so no member can be proven absent.
    """
    return resolve_exhaustive_member_surface_at(source, dot_end_offset, member_ctx)


# -- name-token hits -------------------------------------------------------


@dataclass(frozen=True, slots=True)
class NameTokenHit:
    name: str
    span: Span
    bracketed: bool


def name_token_hit(base: Span, tok: VbaToken, name: str) -> NameTokenHit:
    return NameTokenHit(
        name=name, span=absolute_span(base, tok), bracketed=tok.kind is TokenKind.BRACKETED_IDENTIFIER
    )


def is_bare_or_vba_qualified_intrinsic_call(toks: Sequence[VbaToken], name_index: int) -> bool:
    """Whether the name at `name_index` calls the VBA intrinsic: bare, or qualified
    by `VBA.` alone (`VBA.Left`), never a member of some other object."""
    if name_index < 1 or toks[name_index - 1].raw_text != ".":
        return True
    qualifier = token_name(toks[name_index - 2]) if name_index >= 2 else None
    return (
        qualifier is not None
        and qualifier.lower() == "vba"
        and (name_index < 3 or toks[name_index - 3].raw_text != ".")
    )


def declaration_name_hit(source: str, span: Span, name: str) -> NameTokenHit | None:
    """The first token in the statement span whose name matches `name`."""
    lower = name.lower()
    for tok in statement_tokens(source, span):
        found = token_name(tok)
        if found is not None and found.lower() == lower:
            return name_token_hit(span, tok, found)
    return None


# -- repeated keys ---------------------------------------------------------

_Entry = TypeVar("_Entry")


def report_repeated_keys(
    entries: Iterable[_Entry],
    activity: ConditionalActivityTracker | None,
    key_of: Callable[[_Entry], str | None],
    span_of: Callable[[_Entry], Span],
    report: Callable[[_Entry, _Entry], None],
) -> None:
    """Report every entry that repeats a key an earlier entry already took, once per
    repeat, in source order. `key_of` returns None for an entry the rule does not
    govern.

    Entries in different arms of one `#If` chain never reach the compiler together,
    so they are alternatives rather than repeats however the conditional constants
    evaluate. Asking `activity` for the arm rather than for the activity lets a rule
    keep looking inside a chain it cannot decide instead of going blind to a repeat
    inside one arm (XLIDE issue #58)."""
    taken: dict[str, list[_Entry]] = {}
    for entry in entries:
        key = key_of(entry)
        if key is None:
            continue
        earlier = taken.get(key)
        if earlier is None:
            taken[key] = [entry]
            continue
        hit = next(
            (
                prior
                for prior in earlier
                if activity is None or not activity.mutually_exclusive(span_of(prior), span_of(entry))
            ),
            None,
        )
        earlier.append(entry)
        if hit is not None:
            report(entry, hit)


# -- module-level declaration-statement classifier -------------------------

# Visibility modifiers that may lead a module-level declaration in a procedure
# body (no Static - a Static local is legal inside a procedure).
_PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS: frozenset[str] = frozenset(
    {"public", "private", "friend", "global"}
)

DEFTYPE_KEYWORDS: frozenset[str] = frozenset(
    {
        "defbool", "defbyte", "defcur", "defdate", "defdbl", "defdec", "defint",
        "deflng", "deflnglng", "deflngptr", "defobj", "defsng", "defstr", "defvar",
    }
)


def leading_declaration_modifier_count(toks: Sequence[VbaToken]) -> int:
    i = 0
    while token_text(_at(toks, i)) in _PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS:
        i += 1
    return i


def module_declaration_statement_in_procedure(source: str, span: Span) -> tuple[str, Span] | None:
    """Classify a statement that is really a module-level declaration (Option,
    Attribute, Def*, visibility-led declaration, Type/Enum block) inside a body."""
    toks = statement_tokens_after_leading_label(source, span)
    first = toks[0] if toks else None
    head = token_text(first)
    if first is None:
        return None
    if head == "option":
        return ("Option statements", absolute_span(span, first))
    if head == "attribute":
        return ("Attribute statements", absolute_span(span, first))
    if head in DEFTYPE_KEYWORDS:
        label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " statements"
        return (label, absolute_span(span, first))
    modifier_count = leading_declaration_modifier_count(toks)
    declaration_head = token_text(_at(toks, modifier_count))
    if declaration_head == "type" or declaration_head == "enum":
        tok = _at(toks, modifier_count)
        label = "Type blocks" if declaration_head == "type" else "Enum blocks"
        return (label, absolute_span(span, tok) if tok is not None else absolute_span(span, first))
    if head in _PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS:
        label = (first.canonical_text if first.canonical_text is not None else first.raw_text) + " declarations"
        return (label, absolute_span(span, first))
    return None


# -- undeclared-reference scanner ------------------------------------------
#
# The conservative read-reference scanner the undeclared-variable rule rides
# over. The skip-set is exactly what prevents false positives: it removes every
# token position that is NOT a potential bare variable read (declaration spans,
# line labels, call targets and their argument identifiers, assignment LHS,
# qualified project member/callable qualifiers, New/Is/GoTo/RaiseEvent/AddressOf
# operands, and named-argument labels). Ported line-for-line from shared.ts.


@dataclass(frozen=True, slots=True)
class ValueReadReference:
    name: str
    span: Span
    # `[name]`: in Excel an Evaluate lookup rather than a variable.
    bracketed: bool = False


def for_each_undeclared_reference_span(
    source: str,
    body: Sequence[BodyNode],
    visit: Callable[[Span], None],
    activity: ConditionalActivityTracker | None = None,
) -> None:
    """Visit one span per executable statement / block-header / Do-footer line."""
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if is_leaf_statement(node):
            visit(node.span)
            continue
        # A block: its header and footer lines here, its body next in the walk.
        if isinstance(getattr(node, "body", None), list):
            visit(block_header_line_span(source, node.span))
            if isinstance(node, DoBlockNode):
                footer = block_footer_line_span(source, node.span)
                if footer.start > node.span.start:
                    visit(footer)


def value_read_references(
    source: str,
    span: Span,
    is_known_for_skip: Callable[[str], bool],
    module_signatures: Mapping[str, CallableTypeSignature],
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> list[ValueReadReference]:
    """Bare identifier reads in a statement span, with declaration/call/etc. skipped."""
    toks = statement_tokens(source, span)
    out: list[ValueReadReference] = []
    skip = _undeclared_reference_skip_indexes(
        source, span, toks, is_known_for_skip, module_signatures, project_members
    )
    for i in range(len(toks)):
        if i in skip or not _is_potential_variable_reference_token(_at(toks, i)):
            continue
        # `!` is VBA's default-member accessor, so `rs!CustomerName` and
        # `Forms!frmMain!txtName` are MEMBER ACCESS exactly as a dot is: the name
        # after it belongs to the receiver and was never a variable. The same
        # character is the Single type suffix (`Dim x!`), but a suffix is only ever
        # followed by an operator or the end of the statement, never by a name, so
        # one test covers both.
        if token_raw(_at(toks, i - 1)) in (".", "!"):
            continue
        name = token_name(toks[i])
        if not name:
            continue
        out.append(
            ValueReadReference(
                name=name,
                span=Span(span.start + toks[i].start, span.start + toks[i].end),
                bracketed=toks[i].kind is TokenKind.BRACKETED_IDENTIFIER,
            )
        )
    return out


def token_raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _undeclared_reference_skip_indexes(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    is_known: Callable[[str], bool],
    module_signatures: Mapping[str, CallableTypeSignature],
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> set[int]:
    skip: set[int] = set()
    if len(toks) == 0:
        return skip
    if token_raw(_at(toks, 1)) == ":" or _is_line_label_only_statement(source, span, toks):
        skip.add(0)  # line label declaration
    elif (
        len(toks) == 2
        and is_decimal_line_number(toks[0])
        and _is_line_label_only_statement(source, span, toks[1:])
    ):
        skip.add(1)  # `10 L1:`, a line number and a label (XLIDE issue #230)
    first_executable = first_executable_token_index(toks)
    if module_declaration_statement_in_procedure(source, span):
        for i in range(first_executable, len(toks)):
            skip.add(i)
        return skip
    if token_text(_at(toks, first_executable)) == "implements":
        for i in range(first_executable + 1, len(toks)):
            skip.add(i)
        return skip

    call = call_statement_target(source, span)
    if call is not None:
        call_idx = next(
            (i for i, tok in enumerate(toks) if span.start + tok.start == call.span.start), -1
        )
        if call_idx >= 0:
            skip.add(call_idx)
            if not is_known(call.name):
                # Unknown call targets may be external procedures or unresolved call
                # errors; do not also guess about their argument identifiers.
                for i in range(call_idx + 1, len(toks)):
                    skip.add(i)

    assignment = _simple_assignment_lhs_identifier_index(toks)
    if assignment >= 0:
        skip.add(assignment)

    statement_head = token_text(_at(toks, first_executable))
    for i in range(len(toks)):
        word = token_text(toks[i])
        if _is_qualified_project_callable_qualifier(
            toks, i, module_signatures
        ) or _is_qualified_project_member_qualifier(toks, i, project_members):
            skip.add(i)
        if word == "new" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        if (
            word == "is"
            and _has_earlier_type_of(toks, i)
            and _is_potential_variable_reference_token(_at(toks, i + 1))
        ):
            skip.add(i + 1)
        if _is_label_reference_keyword(word) and _is_potential_variable_reference_token(
            _at(toks, i + 1)
        ):
            skip.add(i + 1)
        # `On n GoTo A, B, C` and `On n GoSub A, B` name a label in every slot of
        # the list, not only the first (XLIDE issue #185).
        if (word == "goto" or word == "gosub") and _is_on_selector_jump(toks, i):
            j = i + 1
            while j < len(toks) and toks[j].raw_text != ":":
                if _is_potential_variable_reference_token(toks[j]):
                    skip.add(j)
                j += 1
        if word == "raiseevent" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        if word == "addressof" and _is_potential_variable_reference_token(_at(toks, i + 1)):
            skip.add(i + 1)
        # A CONTEXTUAL keyword is not a reserved identifier (MS-VBAL 3.3.5.2), so
        # `Dim Text As String` is legal VBA and a reference to `Text` is a real
        # variable read. The scanner therefore accepts one as a name, but each of
        # these words also has a grammar position where it is SYNTAX, and reading
        # the word there would report the statement's own keyword as an undefined
        # variable. Those positions are skipped here.
        if _is_contextual_grammar_word(toks, i, statement_head, first_executable):
            skip.add(i)
        # Open-statement access-clause (MS-VBAL 5.4.5.1.1): in `Open path For
        # mode [Access access] [lock] As #f`, `Access` is a grammar word, not a
        # variable reference. It lexes as an identifier because it is not a
        # reserved identifier (`Dim Access As Long` is legal), so skip it only
        # in its grammar position: inside an Open statement and immediately
        # followed by the access mode `Read` or `Write`. Every other Open
        # grammar word (For/Binary/Input/Output/Append/Random/Read/Write/Lock/
        # Shared/As/Len) already lexes as a keyword and is never scanned.
        if (
            statement_head == "open"
            and word == "access"
            and token_text(_at(toks, i + 1)) in ("read", "write")
        ):
            skip.add(i)
        # `ReDim [Preserve] name(bounds) As TypeName`: the token after `As` is a
        # type reference, not a variable read. Bounds identifiers remain uses.
        # Scoped to ReDim because in other leaf statements the token after `As`
        # IS a genuine value read (`Open path For Output As fnum`, `Name a As b`).
        if (
            statement_head == "redim"
            and word == "as"
            and _is_potential_variable_reference_token(_at(toks, i + 1))
        ):
            skip.add(i + 1)
        if _is_named_argument_label(toks, i):
            skip.add(i)

    return skip


def _is_qualified_project_callable_qualifier(
    toks: Sequence[VbaToken],
    index: int,
    module_signatures: Mapping[str, CallableTypeSignature],
) -> bool:
    if not _is_potential_variable_reference_token(_at(toks, index)) or token_raw(
        _at(toks, index + 1)
    ) != ".":
        return False
    if not _is_potential_variable_reference_token(_at(toks, index + 2)):
        return False
    qualifier = token_name(toks[index])
    member = token_name(toks[index + 2])
    if not qualifier or not member:
        return False
    return qualified_procedure_key(qualifier, member) in module_signatures


def _is_qualified_project_member_qualifier(
    toks: Sequence[VbaToken],
    index: int,
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> bool:
    if (
        not project_members
        or not _is_potential_variable_reference_token(_at(toks, index))
        or token_raw(_at(toks, index + 1)) != "."
    ):
        return False
    if not _is_potential_variable_reference_token(_at(toks, index + 2)):
        return False
    qualifier = token_name(toks[index])
    member = token_name(toks[index + 2])
    if not qualifier or not member:
        return False
    member_lower = member.lower()
    # An ambiguous qualifier is absent from the index, so it is not skipped.
    surface = project_class_members_index(project_members).get(qualifier.lower())
    if surface is None:
        return False
    if surface.kind == "standardModule":
        return True
    # A bare module name in front of a dot is only a legal receiver when the module
    # IS a value: a standard module is a namespace, and documents, forms, and
    # classes marked `VB_PredeclaredId = True` have a default instance. A plain
    # class module is a TYPE, so `Ticket.ChangeTest` is the same `Variable not
    # defined` as any other undeclared name: the VBE refuses to compile it (XLIDE
    # issue #47). Only a VOUCHED-FOR false reports: the attribute is invisible in
    # the code pane, so a host that never read the header leaves this None and the
    # name stays skipped.
    if surface.kind == "class" and surface.predeclared_id is False:
        return False
    return any(candidate.name.lower() == member_lower for candidate in surface.members)


# Words `Open ... For <mode>` accepts that are not reserved identifiers.
_OPEN_MODE_WORDS = frozenset({"binary", "output", "append", "random", "read"})

# Clause keywords an `Open` mode word may follow.
_OPEN_CLAUSE_HEADS = frozenset({"for", "access", "lock"})

# Tokens after which a new statement begins on the same line.
_STATEMENT_OPENERS = frozenset({"then", "else", ":"})


def _is_contextual_grammar_word(
    toks: Sequence[VbaToken], index: int, statement_head: str, first_executable: int
) -> bool:
    """True when a contextual keyword sits in the grammar position that gives it its
    keyword meaning, rather than naming a variable.

    Only the positions reachable inside a PROCEDURE BODY are listed. `Option`,
    `Declare`, and `Property` headers never reach the reference scanner: the first
    two are module-level and an in-procedure declaration is skipped whole.
    """
    word = token_text(_at(toks, index))
    prev = token_text(_at(toks, index - 1))
    # `Exit Property` and the `End Property` footer. Sub, Function, For and Do are
    # reserved, so Property is the only one of these words that reaches the
    # scanner at all.
    if word == "property" and prev in ("exit", "end"):
        return True
    # `On Error GoTo/Resume ...` and the bare `Error <number>` statement. The latter
    # starts a statement, which is not always token 0: `If x Then Error 5 Else Exit
    # Sub` puts it after `Then`.
    if word == "error" and (prev == "on" or index == first_executable or prev in _STATEMENT_OPENERS):
        return True
    # `On Local Error ...` is a form of On Error (XLIDE issue #98): Local is grammar
    # there, and the Error after it is too.
    if word == "local" and prev == "on" and token_text(_at(toks, index + 1)) == "error":
        return True
    if word == "error" and prev == "local" and token_text(_at(toks, index - 2)) == "on":
        return True
    # `Open path For Binary|Output|Append|Random [Access Read] [Lock Read] As #n`.
    # The mode word follows the clause keyword that introduces it; the same words
    # elsewhere in the statement stay readable.
    if statement_head == "open" and word in _OPEN_MODE_WORDS and prev in _OPEN_CLAUSE_HEADS:
        return True
    # `For i = 1 To 10 Step 2`.
    if word == "step" and statement_head == "for":
        return True
    # `As Object` in any statement that reaches here, such as the type clause of a
    # `ReDim ... As Object`.
    return word == "object" and prev == "as"


def _simple_assignment_lhs_identifier_index(toks: Sequence[VbaToken]) -> int:
    start = first_executable_token_index(toks)
    if token_text(_at(toks, start)) == "let" or token_text(_at(toks, start)) == "set":
        start += 1
    eq = top_level_operator_index(list(toks[start:]), "=")
    if eq != 1:
        return -1
    # A name that SPELLS a contextual keyword is still a name, and this skip is what
    # stops the assignment TARGET also being counted as a read (XLIDE #46).
    # Bracketed names are deliberately NOT included: `[If] = x` is Excel's Evaluate
    # shorthand rather than a variable, a separate question this skip must not
    # answer as a side effect.
    name_tok = _at(toks, start)
    return (
        start
        if name_tok is not None
        and (name_tok.kind is TokenKind.IDENTIFIER or _is_non_reserved_keyword(name_tok))
        else -1
    )


def _is_line_label_only_statement(source: str, span: Span, toks: Sequence[VbaToken]) -> bool:
    if len(toks) != 1 or not _is_potential_variable_reference_token(toks[0]):
        return False
    j = span.start + toks[0].end
    while j < len(source) and (source[j] == " " or source[j] == "\t"):
        j += 1
    return j < len(source) and source[j] == ":"


def _is_potential_variable_reference_token(tok: VbaToken | None) -> bool:
    if tok is None:
        return False
    if tok.kind is TokenKind.IDENTIFIER or tok.kind is TokenKind.BRACKETED_IDENTIFIER:
        return True
    return _is_non_reserved_keyword(tok)


def _is_non_reserved_keyword(tok: VbaToken | None) -> bool:
    """A word the VBE capitalizes in its statement context but MS-VBAL 3.3.5.2 does
    not reserve, so `Dim Text As String` and `Dim Error As Long` are both legal and
    a reference to one is a real variable read."""
    if tok is None or tok.kind is not TokenKind.KEYWORD or is_reserved_identifier(tok.raw_text):
        return False
    # `Property` opens a declaration, and the parser already refuses `Dim Property
    # As Long` with invalid-identifier-start, so the word can never reach a body as
    # a variable. Scanning it only adds a second, confusing diagnostic to source
    # that is already reported as malformed.
    return token_text(tok) != "property"


def _has_earlier_type_of(toks: Sequence[VbaToken], before: int) -> bool:
    for i in range(before):
        if token_text(toks[i]) == "typeof":
            return True
    return False


def _is_label_reference_keyword(word: str) -> bool:
    return word == "goto" or word == "gosub" or word == "resume"


def _is_on_selector_jump(toks: Sequence[VbaToken], index: int) -> bool:
    """Whether the GoTo or GoSub at `index` belongs to `On expression GoTo`, not
    `On Error GoTo`."""
    for k in range(index - 1, -1, -1):
        word = token_text(toks[k])
        if word == "on":
            following = token_text(_at(toks, k + 1))
            return following != "error" and following != "local"
        if word == "then" or word == "else" or toks[k].raw_text == ":":
            return False
    return False


def _is_named_argument_label(toks: Sequence[VbaToken], index: int) -> bool:
    if not _is_potential_variable_reference_token(_at(toks, index)):
        return False
    if token_raw(_at(toks, index + 1)) == ":=":
        return True
    return token_raw(_at(toks, index + 1)) == ":" and token_raw(_at(toks, index + 2)) == "="


# -- conditional-compilation branch-order scan -----------------------------


@dataclass(frozen=True, slots=True)
class ConditionalBranchOrderIssue:
    kind: str  # "elseifAfterElse" | "duplicateElse"
    directive: ConditionalDirectiveNode


@dataclass(frozen=True, slots=True)
class ConditionalBranchOrderScan:
    issues: list[ConditionalBranchOrderIssue]
    malformed_block_spans: list[Span]


@dataclass(slots=True)
class _ElseBranchFrame:
    seen_else: bool
    start: Span
    malformed: bool


def scan_conditional_compilation_branch_order(mod: ModuleNode) -> ConditionalBranchOrderScan:
    """Detect #ElseIf-after-#Else and duplicate #Else in conditional blocks, and the
    spans of malformed conditional blocks."""
    stack: list[_ElseBranchFrame] = []
    issues: list[ConditionalBranchOrderIssue] = []
    malformed_block_spans: list[Span] = []
    for occ in collect_conditional_directives(mod):
        directive = occ.directive
        kind = directive.directive_kind
        if kind is ConditionalDirectiveKind.IF:
            stack.append(_ElseBranchFrame(seen_else=False, start=directive.span, malformed=False))
        elif kind is ConditionalDirectiveKind.ELSE_IF:
            frame = stack[-1] if stack else None
            if frame is not None and frame.seen_else:
                frame.malformed = True
                issues.append(ConditionalBranchOrderIssue(kind="elseifAfterElse", directive=directive))
        elif kind is ConditionalDirectiveKind.ELSE:
            frame = stack[-1] if stack else None
            if frame is not None and frame.seen_else:
                frame.malformed = True
                issues.append(ConditionalBranchOrderIssue(kind="duplicateElse", directive=directive))
            if frame is not None:
                frame.seen_else = True
        elif kind is ConditionalDirectiveKind.END_IF:
            popped = stack.pop() if stack else None
            if popped is not None and popped.malformed:
                malformed_block_spans.append(Span(popped.start.start, directive.span.end))
        # Const, Unknown: no branch-order effect.
    return ConditionalBranchOrderScan(issues=issues, malformed_block_spans=malformed_block_spans)


# -- names a statement or a procedure mentions --------------------------------


def names_in(source: str, span: Span) -> set[str]:
    """The lowercased names a statement mentions, which a block containing it may
    change (XLIDE issue #237)."""
    out: set[str] = set()
    for tok in statement_tokens_after_leading_label(source, span):
        name = token_name(tok)
        if name:
            out.add(name.lower())
    return out


class _HasBody(Protocol):
    @property
    def body(self) -> Sequence[BodyNode]: ...


def name_mentions(
    source: str, procedure: _HasBody, activity: ConditionalActivityTracker | None
) -> dict[str, int]:
    """How many times each name appears in the procedure's active code: its
    statements, and the header and footer lines of its blocks, where
    `For Each v In c` assigns v."""
    procedure_body = procedure.body
    out: dict[str, int] = {}

    def count(span: Span) -> None:
        for tok in statement_tokens(source, span):
            name = token_name(tok)
            if name:
                lower = name.lower()
                out[lower] = out.get(lower, 0) + 1

    for node in iter_body_nodes(procedure_body, inactive_node_skip(activity)):
        if is_leaf_statement(node):
            count(node.span)
        elif isinstance(getattr(node, "body", None), list):
            count(block_header_line_span(source, node.span))
            count(block_footer_line_span(source, node.span))
    return out


# -- For Each's source and the array bound reads -------------------------------

# Words that open no variable, member or call.
_NON_SOURCE_WORDS: frozenset[str] = frozenset({"null", "true", "false", "nothing", "empty", "new", "not", "typeof"})

# Words that join two operands.
_OPERATOR_WORDS: frozenset[str] = frozenset({"and", "or", "xor", "eqv", "imp", "mod", "like", "is"})

_LITERAL_KINDS = frozenset(
    {TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL, TokenKind.STRING_LITERAL, TokenKind.DATE_LITERAL}
)


def source_expression_syntax_problem(source_expression: str) -> str | None:
    """Why an expression cannot stand where the grammar wants a variable, a member
    chain or a call, or None: For Each's source (XLIDE issue #239) and the array
    UBound and LBound read, each measured in Excel 16.0. `In 5`, `In "abc"`,
    `In (c)`, `In New Collection`, `In -v`, `In v & v`, `In Len("abc")`, Len being a
    special form, and `UBound(5)` are each "Syntax error". `In Split("a" & "b")`,
    `In [A1:B2]` and `UBound(Array(1))` compile."""
    # Lexed alone, a leading '#' reads as a directive: `#1/1/2000#`.
    text = js_trim(source_expression)
    if _DATE_LITERAL_TEXT_RE.fullmatch(text):
        return f"the literal {text}"
    significant = [
        tok
        for tok in raw_expression_tokens(text)
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    if not significant:
        return None
    first = significant[0]
    first_word = token_text(first)
    if first.kind in _LITERAL_KINDS:
        return f"the literal {first.raw_text}"
    if first.raw_text in ("(", "-", "+"):
        return f"an expression that opens with '{first.raw_text}'"
    if first.kind is TokenKind.KEYWORD and (
        first_word in _NON_SOURCE_WORDS
        or (first_word == "len" and len(significant) > 1 and significant[1].raw_text == "(")
    ):
        return f"'{first.raw_text}'"
    depth = 0
    for tok in significant[1:]:
        if tok.raw_text in ("(", "["):
            depth += 1
        elif tok.raw_text in (")", "]"):
            depth -= 1
        elif depth == 0 and (
            (tok.kind is TokenKind.OPERATOR and tok.raw_text != "!")
            or (tok.kind is TokenKind.KEYWORD and token_text(tok) in _OPERATOR_WORDS)
        ):
            return f"an expression joined by '{tok.raw_text}'"
    return None


# /^#[^#]*#$/
_DATE_LITERAL_TEXT_RE = re.compile(r"#[^#]*#")


# -- loops that may be left -----------------------------------------------------

_LOOP_LEAVES: frozenset[str] = frozenset({"for", "sub", "function", "property"})


def body_may_leave_loop(source: str, body: Sequence[BodyNode]) -> bool:
    """True when a statement in the loop's body can leave the loop before its last
    pass: `Exit For` (not one belonging to a nested For), `Exit Sub`/`Function`/
    `Property`, `GoTo`, or `End` (XLIDE issues #145, #356). Walked on an explicit
    stack; the answer does not depend on the order statements are seen in."""
    stack: list[tuple[Sequence[BodyNode], bool]] = [(body, False)]
    while stack:
        nodes, inside_nested_for = stack.pop()
        for node in nodes:
            if is_leaf_statement(node):
                toks = statement_tokens_after_leading_label(source, node.span)
                for i, tok in enumerate(toks):
                    word = token_text(tok)
                    if word == "goto" or (word == "end" and len(toks) == 1):
                        return True
                    if word == "exit":
                        target = token_text(_at(toks, i + 1))
                        if target in _LOOP_LEAVES and (target != "for" or not inside_nested_for):
                            return True
            else:
                inner = getattr(node, "body", None)
                if isinstance(inner, list):
                    stack.append((inner, inside_nested_for or isinstance(node, ForBlockNode)))
    return False


# -- built-ins that read their argument's value ----------------------------------

# VBA built-ins that read their argument's value: on an object still Nothing that
# raises 91 (XLIDE issue #415, measured in Excel 16.0 with CStr and Len).
ONE_VALUE_BUILTINS: frozenset[str] = frozenset(
    {"cstr", "len", "lenb", "clng", "cint", "cdbl", "csng", "ccur", "cbool", "cdate", "val", "trim", "ucase", "lcase"}
)


def builtin_name_before(toks: Sequence[VbaToken], index: int) -> int:
    """The index of the name whose argument list holds toks[index], `$` spellings
    included, or -1."""
    depth = 0
    for i in range(index - 1, -1, -1):
        raw = toks[i].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            if depth == 0:
                at = i - 2 if token_raw(_at(toks, i - 1)) == "$" else i - 1
                if token_raw(_at(toks, at - 1)) == "." and token_text(_at(toks, at - 2)) != "vba":
                    return -1
                return at
            depth -= 1
    return -1


__all__ = [
    "DEFTYPE_KEYWORDS",
    "ONE_VALUE_BUILTINS",
    "ConditionalBranchOrderIssue",
    "ConditionalBranchOrderScan",
    "NameTokenHit",
    "ValueReadReference",
    "bankers_round",
    "body_may_leave_loop",
    "builtin_name_before",
    "declaration_name_hit",
    "for_each_undeclared_reference_span",
    "is_bare_or_vba_qualified_intrinsic_call",
    "leading_declaration_modifier_count",
    "module_declaration_statement_in_procedure",
    "name_mentions",
    "name_token_hit",
    "names_in",
    "report_repeated_keys",
    "resolve_exhaustive_member_surface",
    "scan_conditional_compilation_branch_order",
    "source_expression_syntax_problem",
    "token_raw",
    "value_read_references",
]

