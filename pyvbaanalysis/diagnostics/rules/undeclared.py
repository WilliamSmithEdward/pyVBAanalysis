"""Rule family: unresolved-name rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/undeclared.ts. Four rules:
Option Explicit presence (style), undeclared variable reads/writes, and unknown /
non-callable bare call statements.

Self-gating preserves the no-false-positive guarantee: `check_undeclared_variables`
and `check_unknown_call_statement` no-op unless the caller supplies the project
identifier/procedure sets (the cross-module surface). The member-not-found rule
(`checkMemberNotFound` in the TS) needs the host member-completion surface and is
deferred to M9; only its pure helper `member_access_references` is ported here.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...call.call_context import bare_call_statement_target as call_statement_target
from ...conditional import ConditionalActivityTracker
from ...host import (
    application_member_names,
    resolve_host_constant,
    resolve_host_global,
)
from ...lexer.keyword_table import is_reserved_identifier
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize
from ...parser.nodes import (
    DeclareNode,
    EnumNode,
    LeafStatementNode,
    ModuleNode,
    OptionNode,
    ProcedureNode,
    Span,
    TypeNode,
    VariableGroupNode,
)
from ...runtime import (
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
)
from ...symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolutionScope,
)
from ...symbols.symbol_model import (
    ModuleSymbols,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
)
from ..call_extraction import (
    CallableTypeSignature,
    CallArguments,
    extract_call,
    is_named_slot,
)
from ..callable_signatures import (
    callable_type_signatures_for,
    is_non_callable_symbol,
)
from ..context import PushFn, statement_tokens
from ..model import VbaCreateProcedureStubData, VbaDiagnosticData, VbaEdit
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    set_assignment_target,
    token_name,
)
from ...completion import MemberCompletionContext
from .shared import (
    ValueReadReference,
    for_each_undeclared_reference_span,
    resolve_exhaustive_member_surface,
    value_read_references,
)
from ...types.type_inference import (
    procedure_symbol_for,
    source_identifier_binding,
    source_identifier_bound,
)

_VBA_IDENTIFIER_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_ENDS_WITH_BLANK_PHYSICAL_LINE_RE = re.compile(r"(?:\r\n|\r|\n)[ \t]*(?:\r\n|\r|\n)$")


# -- member-access references (pure helper; checkMemberNotFound deferred to M9) --


@dataclass(frozen=True, slots=True)
class MemberAccessReference:
    member: str
    member_span: Span
    dot_end_offset: int


def member_access_references(source: str, span: Span) -> list[MemberAccessReference]:
    """The `.member` references in a statement span (member after a dot)."""
    toks = statement_tokens(source, span)
    out: list[MemberAccessReference] = []
    for i in range(len(toks) - 1):
        if toks[i].raw_text != ".":
            continue
        member = token_name(toks[i + 1])
        if not member:
            continue
        out.append(
            MemberAccessReference(
                member=member,
                member_span=Span(span.start + toks[i + 1].start, span.start + toks[i + 1].end),
                dot_end_offset=span.start + toks[i].end,
            )
        )
    return out


def check_member_not_found(
    source: str,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """`receiver.Member` where the receiver type resolves to an EXHAUSTIVE member
    surface and `Member` is genuinely absent: "Method or data member not found".

    Ported from checkMemberNotFound (undeclared.ts). Rides the shared
    procedure-statement walk. The no-false-positive contract is the exhaustive
    gate in resolve_exhaustive_member_surface: a non-exhaustive host type, an
    Object/Variant receiver, or an unresolved receiver yields no surface, so the
    rule stays silent. Public fields and known members are present in the surface,
    so they never fire. Host events are excluded from object surfaces, so an event
    name on an exhaustive receiver is reported absent (matching VBE)."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        def visitor(stmt: LeafStatementNode) -> None:
            for ref in member_access_references(source, stmt.span):
                surface = resolve_exhaustive_member_surface(
                    source, ref.dot_end_offset, member_ctx
                )
                if surface is None:
                    continue
                lower = ref.member.lower()
                if any(candidate.name.lower() == lower for candidate in surface.members):
                    continue
                push(
                    "memberNotFound",
                    f"Method or data member not found: '{surface.owner}.{ref.member}'.",
                    ref.member_span,
                )

        return visitor

    return factory


# -- unknownCallStatement --------------------------------------------------


def check_unknown_call_statement(
    source: str,
    symbols: ModuleSymbols,
    known_procedures: AbstractSet[str],
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """A bare call statement whose callee resolves to nothing: "Sub or Function not
    defined". Resolution covers project procedures, source bindings, Application
    members, host globals, and the VBA runtime, so only truly-unknown names fire."""
    known = {name.lower() for name in known_procedures}
    app_members = application_member_names()

    def is_known(name: str, proc_sym: VbaSymbol | None) -> bool:
        lower = name.lower()
        return (
            lower in known
            or source_identifier_bound(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.CALL
            )
            or lower in app_members
            or resolve_host_global(name) is not None
            or resolve_runtime_object(name) is not None
            or resolve_runtime_function(name) is not None
        )

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            hit = call_statement_target(source, stmt.span)
            if hit is not None and not is_known(hit.name, proc_sym):
                call = extract_call(source, stmt.span)
                data = (
                    _create_procedure_stub_data(source, call)
                    if call is not None
                    and call.name_span.start == hit.span.start
                    and call.name_span.end == hit.span.end
                    else None
                )
                push(
                    "unknownCallStatement",
                    f"Sub or Function not defined: '{hit.name}'.",
                    hit.span,
                    data,
                )

        return visitor

    return factory


def _create_procedure_stub_data(source: str, call: CallArguments) -> VbaDiagnosticData | None:
    if not _is_generated_stub_identifier(call.name):
        return None
    params = _generated_stub_parameters(call)
    if params is None:
        return None
    eol = _detect_eol(source)
    if len(source) == 0:
        leading = ""
    else:
        first = "" if source.endswith("\n") or source.endswith("\r") else eol
        second = "" if _ends_with_blank_physical_line(source) else eol
        leading = f"{first}{second}"
    text = f"{leading}Private Sub {call.name}({', '.join(params)}){eol}End Sub{eol}"
    return VbaDiagnosticData(
        create_procedure_stub=VbaCreateProcedureStubData(
            procedure_name=call.name,
            edit=VbaEdit(span=Span(len(source), len(source)), new_text=text),
        )
    )


def _generated_stub_parameters(call: CallArguments) -> list[str] | None:
    if any(len(slot) == 0 for slot in call.slots):
        return None
    named = [is_named_slot(slot) for slot in call.slots]
    if any(named) and not all(named):
        return None
    used: set[str] = set()
    params: list[str] = []
    for i in range(len(call.slots)):
        name = (
            _generated_named_argument_parameter_name(call.slots[i])
            if named[i]
            else f"arg{i + 1}"
        )
        if not name or name.lower() in used:
            return None
        used.add(name.lower())
        params.append(f"ByVal {name} As Variant")
    return params


def _generated_named_argument_parameter_name(slot: Sequence[VbaToken]) -> str | None:
    raw = slot[0].raw_text if slot else None
    if not raw or raw.startswith("["):
        return None
    return raw if _is_generated_stub_identifier(raw) else None


def _is_generated_stub_identifier(name: str) -> bool:
    return _VBA_IDENTIFIER_NAME_RE.match(name) is not None and not is_reserved_identifier(name)


def _detect_eol(source: str) -> str:
    return "\r\n" if "\r\n" in source else "\n"


def _ends_with_blank_physical_line(source: str) -> bool:
    return _ENDS_WITH_BLANK_PHYSICAL_LINE_RE.search(source) is not None


# -- nonCallableCallStatement ----------------------------------------------


def check_non_callable_call_statement(
    source: str,
    symbols: ModuleSymbols,
    known_procedures: AbstractSet[str] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """A call statement whose callee resolves to a non-callable declaration (a
    variable, constant, enum, or Type) is a compile error."""
    known = (
        None if known_procedures is None else {name.lower() for name in known_procedures}
    )

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            call = extract_call(source, stmt.span)
            if call is None:
                return
            binding = source_identifier_binding(
                symbols, proc_sym, project_visible_symbols, call.name, BareIdentifierContext.CALL
            )
            if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS:
                return
            if (
                binding.tier is BareIdentifierResolutionScope.PROJECT
                and known is not None
                and call.name.lower() in known
            ):
                return
            target = next(
                (symbol for symbol in binding.definitions if is_non_callable_symbol(symbol)), None
            )
            if target is None:
                return
            if _call_target_feeds_member_access(source, stmt.span, call):
                return
            push(
                "nonCallableCallStatement",
                f"Cannot call '{call.name}' because it resolves to "
                f"{_symbol_kind_label(target)}, not a Sub or Function.",
                call.name_span,
            )

        return visitor

    return factory


def _call_target_feeds_member_access(source: str, span: Span, call: CallArguments) -> bool:
    toks = [
        t
        for t in tokenize(source[span.start : span.end])
        if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE
    ]
    rel_callee_start = call.name_span.start - span.start
    callee_idx = next((i for i, t in enumerate(toks) if t.start == rel_callee_start), -1)
    after = toks[callee_idx + 1] if 0 <= callee_idx and callee_idx + 1 < len(toks) else None
    if callee_idx < 0 or after is None or after.raw_text != "(":
        return False
    close = match_paren_from(toks, callee_idx + 1)
    next_after = toks[close + 1] if 0 <= close and close + 1 < len(toks) else None
    return close >= 0 and next_after is not None and next_after.raw_text == "."


def _symbol_kind_label(sym: VbaSymbol) -> str:
    if sym.kind is VbaSymbolKind.PARAMETER:
        return "a parameter"
    if sym.kind is VbaSymbolKind.LOCAL_VARIABLE:
        return "a local variable"
    if sym.kind is VbaSymbolKind.MODULE_VARIABLE:
        return "a module variable"
    if sym.kind is VbaSymbolKind.CONSTANT:
        return "a constant"
    if sym.kind is VbaSymbolKind.ENUM:
        return "an enum type"
    if sym.kind is VbaSymbolKind.ENUM_MEMBER:
        return "an enum member"
    if sym.kind is VbaSymbolKind.TYPE:
        return "a user-defined type"
    return "a non-callable declaration"


# -- optionExplicit --------------------------------------------------------

_EXPLICIT_OPTION_RE = re.compile(r"^explicit\b", re.IGNORECASE)


def check_option_explicit(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """A code module with real code but no Option Explicit lets variables be used
    without declaration. Empty/attribute-only modules are skipped (no noise)."""
    has_explicit = False
    has_code = False
    for member in active_module_members(mod, activity):
        if isinstance(member, OptionNode) and _EXPLICIT_OPTION_RE.match(member.option_text.strip()):
            has_explicit = True
        if _is_code_member(member):
            has_code = True
    if has_explicit or not has_code:
        return
    push(
        "optionExplicitMissing",
        'Option Explicit is not specified; variables can be used without being '
        'declared. Add "Option Explicit" to the top of the module.',
        Span(0, 0),
    )


def _is_code_member(member: object) -> bool:
    return isinstance(member, (ProcedureNode, VariableGroupNode, TypeNode, EnumNode, DeclareNode))


# -- undeclaredVariables ---------------------------------------------------


def check_undeclared_variables(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    known_identifiers: AbstractSet[str] | None,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_members: Sequence[VbaProjectClassMembers] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> None:
    """With Option Explicit, a variable must be declared before it is assigned or
    read. Self-gated on the caller supplying the project-visible identifier set, so
    cross-module globals and enum members never false-positive."""
    if not _has_option_explicit(mod, activity) or known_identifiers is None:
        return

    known = {name.lower() for name in known_identifiers}
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    app_members = application_member_names()

    def is_known(
        name: str, proc_sym: VbaSymbol | None, context: BareIdentifierContext
    ) -> bool:
        lower = name.lower()
        return (
            lower == "vba"
            or source_identifier_bound(symbols, proc_sym, project_visible_symbols, name, context)
            or lower in known
            or lower in app_members
            or resolve_host_global(name) is not None
            or resolve_host_constant(name) is not None
            or resolve_runtime_constant(name) is not None
            or resolve_runtime_object(name) is not None
            or resolve_runtime_function(name) is not None
        )

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)

        def visit(span: Span, proc_sym: VbaSymbol | None = proc_sym) -> None:
            reported: set[str] = set()

            def report(
                name: str, ref_span: Span, mode: str, context: BareIdentifierContext
            ) -> None:
                key = f"{ref_span.start}:{ref_span.end}"
                if key in reported or is_known(name, proc_sym, context):
                    return
                reported.add(key)
                push(
                    "undeclaredVariable",
                    f"Variable not defined: '{name}'. Declare it before {mode}, "
                    f"or remove Option Explicit.",
                    ref_span,
                )

            scalar_target = bare_assignment_target(source, span)
            object_target = None if scalar_target is not None else set_assignment_target(source, span)
            target = scalar_target if scalar_target is not None else object_target
            if target is not None:
                report(
                    target[0], target[1], "assigning to it", BareIdentifierContext.ASSIGNMENT_TARGET
                )
            for ref in _undeclared_read_references(
                source,
                span,
                lambda name: is_known(name, proc_sym, BareIdentifierContext.EXPRESSION),
                module_signatures,
                project_members,
            ):
                report(ref.name, ref.span, "using it", BareIdentifierContext.EXPRESSION)

        for_each_undeclared_reference_span(source, member.body, visit, activity)


def _undeclared_read_references(
    source: str,
    span: Span,
    is_known: Callable[[str], bool],
    module_signatures: Mapping[str, CallableTypeSignature],
    project_members: Sequence[VbaProjectClassMembers] | None,
) -> list[ValueReadReference]:
    return [
        ref
        for ref in value_read_references(source, span, is_known, module_signatures, project_members)
        if not is_known(ref.name)
    ]


def _has_option_explicit(
    mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> bool:
    return any(
        isinstance(member, OptionNode) and _EXPLICIT_OPTION_RE.match(member.option_text.strip())
        for member in active_module_members(mod, activity)
    )
