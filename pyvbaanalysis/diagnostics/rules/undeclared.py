"""Rule family: unresolved-name rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/undeclared.ts. Rules:
Option Explicit presence (style), undeclared variable reads/writes, VBA library
procedures named bare where a value goes, unknown / non-callable bare call
statements, and member-not-found.

Self-gating preserves the no-false-positive guarantee: `check_undeclared_variables`
and `check_unknown_call_statement` no-op unless the caller supplies the project
identifier/procedure sets (the cross-module surface). The member-not-found rule
(`check_member_not_found`) rides the host member-completion surface, gated on the
exhaustive-surface check; its pure helper `member_access_references` collects the
`.member` references it inspects.

Port-only deviation: an undeclared-variable diagnostic carries no `declareVariable`
data (upstream's declarationDataFor). That editor quick fix types the assigned
value with resolveExpressionType, which the port does not have; the diagnostic
itself (code, message, span) is upstream's.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...call.call_context import bare_call_statement_target as call_statement_target
from ...completion.member_access import (
    private_member_owner_at,
    project_class_member_at,
    project_type_at,
)
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...host import (
    application_member_names,
    resolve_host_constant,
    resolve_host_global,
)
from ...host.host_libraries import HOST_LIBRARY_NAMES
from ...host.host_model import (
    HostObjectModel,
    get_host_members,
    resolve_host_enum,
    resolve_host_global_member,
)
from ...host.ms_forms_form_control_members import MSFORMS_FORM_CONTROL_MEMBERS
from ...identity_cache import IdentityLru
from ...js_compat import JS_WHITESPACE, js_trim
from ...lexer.keyword_table import is_reserved_identifier
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    DeclareNode,
    EnumNode,
    LeafStatementNode,
    ModuleNode,
    OptionNode,
    ProcedureNode,
    Span,
    TypeNode,
    VariableGroupNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...runtime import (
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
    resolve_vba_library_qualifier,
)
from ...runtime.vba_runtime import VbaRuntimeFunction
from ...symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolutionScope,
)
from ...symbols.symbol_model import (
    ImplicitMember,
    ModuleSymbolKind,
    ModuleSymbols,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
)
from ...types.type_names import is_known_scalar_type, normalize_type
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
    first_executable_token_index,
    for_each_statement,
    for_each_variable_group,
    set_assignment_target,
    statement_and_branch_spans,
    token_name,
    token_text,
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


# -- builtinsReadBare ------------------------------------------------------

_NAMES_ON_TOKENS_AFTER = frozenset({"(", ".", "!", ":=", "$"})


def check_builtins_read_bare(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    module_kind: ModuleSymbolKind | None,
    host_model: HostObjectModel | None,
    designer_class: str | None,
    implicit_members: Sequence[ImplicitMember] | None,
    push: PushFn,
    own_members: AbstractSet[str] = frozenset(),
) -> None:
    """A VBA library procedure named bare where a value goes (XLIDE issue #318,
    measured in Excel 16.0, with or without Option Explicit). One that needs an
    argument, `Main = Left` or `TypeName(Kill)`, is "Argument not optional"; a
    statement that takes none, `Main = Beep` or `Reset`, is "Expected Function or
    variable". One whose arguments are all optional, Now or Timer, gives its value.
    A name the module, the project, the host or the module's own object declares is
    theirs."""
    if module_kind is ModuleSymbolKind.USERFORM and implicit_members is None:
        return
    app_members = application_member_names(host_model)
    designer_members = designer_class_member_names(designer_class, host_model)
    implicit_names = {member.name.lower() for member in implicit_members or ()}
    explicit = _has_option_explicit(mod, activity)

    def builtin(name: str, proc_sym: VbaSymbol | None) -> VbaRuntimeFunction | None:
        lower = name.lower()
        if (
            lower in app_members
            or lower in designer_members
            or lower in own_members
            or lower in implicit_names
            or source_identifier_bound(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )
            or resolve_host_global(name, host_model) is not None
            or resolve_host_global_member(name, host_model) is not None
            or resolve_host_constant(name, host_model) is not None
            or resolve_host_enum(name, host_model) is not None
            or resolve_runtime_constant(name) is not None
            or resolve_runtime_object(name) is not None
        ):
            return None
        return resolve_runtime_function(name)

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        redim = _redim_target_names_in(source, member.body, activity)

        def visit(
            stmt: LeafStatementNode,
            proc_sym: VbaSymbol | None = proc_sym,
            redim: AbstractSet[str] = redim,
        ) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                assignment = bare_assignment_target(source, span)
                value_from = (
                    next(
                        (i for i, tok in enumerate(toks) if tok.start == assignment[2][0].start),
                        -1,
                    )
                    if assignment is not None and len(assignment[2]) > 0
                    else -1
                )
                depth = 0
                for i, tok in enumerate(toks):
                    depth += 1 if tok.raw_text == "(" else -1 if tok.raw_text == ")" else 0
                    next_raw = _raw_at(toks, i + 1)
                    # A name after AddressOf is addressof-misuse's (XLIDE issue #299).
                    if (
                        tok.kind is not TokenKind.IDENTIFIER
                        or (depth == 0 and (value_from < 0 or i < value_from))
                        or (_raw_at(toks, i - 1) or "") in (".", "!")
                        or token_text(_at(toks, i - 1)) == "addressof"
                        or (next_raw or "") in _NAMES_ON_TOKENS_AFTER
                        or tok.raw_text.lower() in redim
                    ):
                        continue
                    runtime = builtin(tok.raw_text, proc_sym)
                    if runtime is None:
                        continue
                    at = Span(span.start + tok.start, span.start + tok.end)
                    # `Line` and `Name` open statements and name no procedure: read
                    # as a value under Option Explicit, each is "Variable not
                    # defined".
                    if runtime.name == "Line" or runtime.name == "Name":
                        if explicit:
                            push(
                                "undeclaredVariable",
                                f"Variable not defined: '{tok.raw_text}'. It opens the "
                                f"{runtime.name} statement, which gives no value. Declare a "
                                "variable of that name, or remove Option Explicit.",
                                at,
                            )
                        continue
                    required = _runtime_required_count(runtime)
                    if required > 0:
                        needs = "an argument" if required == 1 else f"{required} arguments"
                        push(
                            "argumentCount",
                            f"'{tok.raw_text}' needs {needs}, and is named here with none where "
                            "a value goes. This is a VBE compile error: Argument not optional.",
                            at,
                        )
                    elif runtime.kind == "statement":
                        push(
                            "subUsedAsValue",
                            f"'{tok.raw_text}' is a statement, which returns nothing, so it "
                            "cannot be used as a value. This is a VBE compile error: Expected "
                            "Function or variable.",
                            at,
                        )

        for_each_statement(member.body, visit, activity)


_PARAM_ARRAY_OR_OPTIONAL_RE = re.compile(r"^(ParamArray|Optional)\b", re.IGNORECASE | re.ASCII)


def _runtime_required_count(runtime: VbaRuntimeFunction) -> int:
    """How many arguments a library procedure needs: those not in brackets, from its
    signature."""
    if runtime.params is not None:
        return sum(1 for param in runtime.params if not param.optional and not param.param_array)
    open_ = runtime.signature.find("(")
    # JS slice: an end of -1 (no `)`) counts from the end, as Python's does.
    parts = (
        runtime.signature[open_ + 1 : runtime.signature.find(")", open_)]
        if open_ >= 0
        else runtime.signature[len(runtime.name) :]
    )
    return sum(
        1
        for part in (js_trim(one) for one in parts.split(","))
        if part != "" and not part.startswith("[") and _PARAM_ARRAY_OR_OPTIONAL_RE.match(part) is None
    )


# -- member-access references (pure helper for checkMemberNotFound) --


@dataclass(frozen=True, slots=True)
class MemberAccessReference:
    member: str
    member_span: Span
    dot_end_offset: int
    # The statement's tokens, and where the member is among them.
    toks: Sequence[VbaToken] = ()
    index: int = 0


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
                toks=toks,
                index=i + 1,
            )
        )
    return out


def _project_member_form_problem(
    source: str, ref: MemberAccessReference, member_ctx: MemberCompletionContext
) -> str | None:
    """A class member used in a form the VBE refuses while compiling (XLIDE issue
    #224, measured in Excel 16.0):

    - A property whose Get takes a required index, used without one: `Main = c.Idx`,
      `c.Idx = 5`. "Argument not optional".
    - A Public field of a value type given arguments: `c.Field(1)` is "Wrong number
      of arguments or invalid property assignment", and as the target of a Let,
      `c.Field(1) = 5`, "Can't assign to read-only property". A Variant, Collection
      or Object field takes them, and `c.Field()` compiles.
    """
    next_raw = _raw_at(ref.toks, ref.index + 1)
    if next_raw == "." or token_text(_at(ref.toks, 0)) == "set":
        return None
    member = project_class_member_at(source, ref.dot_end_offset, ref.member, member_ctx)
    if member is None or member.kind != "property":
        return None
    if member.signature is not None:
        open_ = member.signature.find("(")
        first_param = member.signature[open_ + 1 :].lstrip(JS_WHITESPACE) if open_ >= 0 else ""
        index_required = (
            len(first_param) > 0 and not first_param.startswith(")") and not first_param.startswith("[")
        )
        return (
            f"Argument not optional: property '{member.name}' takes an index, as in "
            f"{member.signature}. This is a VBE compile error."
            if index_required and next_raw != "("
            else None
        )
    field = member.writable is True and not member.let_accessor and not member.set_accessor
    type_ = normalize_type(member.returns)
    if (
        not field
        or type_ is None
        or type_ == "variant"
        or not is_known_scalar_type(type_)
        or next_raw != "("
    ):
        return None
    close = match_paren_from(ref.toks, ref.index + 1)
    if close != ref.index + 2 and close > 0:
        assigned = _raw_at(ref.toks, close + 1) == "=" and ref.index == (
            1 if _raw_at(ref.toks, 0) == "." else 2
        )
        return (
            f"'{member.name}' is a field of type {member.returns}, which takes no index, so "
            f"'{member.name}(...)' is no place to assign. This is a VBE compile error: Can't "
            "assign to read-only property."
            if assigned
            else f"'{member.name}' is a field of type {member.returns}, which takes no arguments. "
            "This is a VBE compile error: Wrong number of arguments or invalid property "
            "assignment."
        )
    return None


_MSFORMS_PREFIX_RE = re.compile(r"^MSForms\.", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class _FormControl:
    name: str
    type: str


def _form_control_without_member(
    source: str,
    ref: MemberAccessReference,
    member_ctx: MemberCompletionContext,
    own_names: AbstractSet[str],
) -> _FormControl | None:
    """A form's control reached through the form, with a member its class lacks:
    `f.T1.Nope`, `Me.T1.Nope`, a bare `T1.Nope` inside the form. The VBE binds those
    while compiling for the classes in MSFORMS_FORM_CONTROL_MEMBERS; a variable
    declared As MSForms.TextBox, and a Frame or an OptionButton on the form, it
    leaves to run time (XLIDE issue #226, measured in Excel 16.0)."""
    control_token = _at(ref.toks, ref.index - 2)
    name = token_name(control_token) if control_token is not None else None
    if not name or _raw_at(ref.toks, ref.index - 1) != ".":
        return None
    lower = name.lower()
    type_: str | None
    if _raw_at(ref.toks, ref.index - 3) == ".":
        # `f.T1.Nope`: T1 must be a control of the form the receiver is.
        form = project_type_at(
            source,
            ref.dot_end_offset - (ref.toks[ref.index - 1].end - ref.toks[ref.index - 3].end),
            member_ctx,
        )
        if form is None or form.kind != "userform" or form.exhaustive is not True:
            return None
        type_ = next(
            (
                member.returns
                for member in form.members
                if member.name.lower() == lower
                and _MSFORMS_PREFIX_RE.match(member.returns or "") is not None
            ),
            None,
        )
    else:
        # A bare `T1.Nope` inside the form, where no local or parameter takes the
        # name.
        me_type = member_ctx.me_project_type.lower() if member_ctx.me_project_type is not None else None
        self_ = next(
            (
                candidate
                for candidate in member_ctx.project_class_members or ()
                if candidate.kind == "userform"
                and candidate.exhaustive is True
                and candidate.name.lower() == me_type
            ),
            None,
        )
        type_ = (
            None
            if lower in own_names or self_ is None
            else next(
                (
                    member.returns
                    for member in self_.members
                    if member.name.lower() == lower
                    and _MSFORMS_PREFIX_RE.match(member.returns or "") is not None
                ),
                None,
            )
        )
    if not type_:
        return None
    members = MSFORMS_FORM_CONTROL_MEMBERS.get(type_)
    if members is None or any(member.lower() == ref.member.lower() for member in members):
        return None
    return _FormControl(name, type_)


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


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
        # Names a bare control reference would lose to: the procedure's parameters
        # and locals.
        own_names = {param.name.lower() for param in member.params}

        def note_group(group: VariableGroupNode) -> None:
            for decl in group.declarations:
                own_names.add(decl.name.lower())

        for_each_variable_group(member.body, note_group)

        def visitor(stmt: LeafStatementNode) -> None:
            for ref in member_access_references(source, stmt.span):
                surface = resolve_exhaustive_member_surface(
                    source, ref.dot_end_offset, member_ctx
                )
                if surface is None or surface.has_member(ref.member):
                    form = _project_member_form_problem(source, ref, member_ctx)
                    if form is not None:
                        push("argumentCount", form, ref.member_span)
                        continue
                    control = _form_control_without_member(source, ref, member_ctx, own_names)
                    if control is not None:
                        control_type = (
                            control.type[len("MSForms.") :]
                            if control.type.startswith("MSForms.")
                            else control.type
                        )
                        push(
                            "memberNotFound",
                            f"Method or data member not found: '{control.name}.{ref.member}'. "
                            f"The form's {control_type} has no member of that name.",
                            ref.member_span,
                        )
                        continue
                    owner = private_member_owner_at(
                        source, ref.dot_end_offset, ref.member, member_ctx
                    )
                    if owner:
                        push(
                            "memberNotFound",
                            f"Method or data member not found: '{owner}.{ref.member}'. It is "
                            f"Private to {owner}, and no reference through an object reaches a "
                            "Private member, Me included.",
                            ref.member_span,
                        )
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
    host_model: HostObjectModel | None,
    designer_class: str | None,
    push: PushFn,
    project_types: Sequence[VbaProjectClassMembers] | None = None,
    own_members: AbstractSet[str] = frozenset(),
) -> ProcedureStatementVisitor:
    """A bare call statement whose callee resolves to nothing: "Sub or Function not
    defined". Resolution covers project procedures, source bindings, Application
    members, host globals, and the VBA runtime, so only truly-unknown names fire."""
    known = {name.lower() for name in known_procedures}
    # The host injects Application's members into the global scope, so a bare call
    # may legitimately bind to one of them (Calculate, Volatile, ...).
    app_members = application_member_names(host_model)
    # A module IS its designer's class, so that class's own methods are in scope
    # unqualified: Requery in an Access form.
    designer_members = designer_class_member_names(designer_class, host_model)

    # An Event is no procedure: `Done` or `Call Done(1)` in the class that declares
    # only the Event is "Sub or Function not defined" (XLIDE issue #266, measured in
    # Excel 16.0). A Sub of the same name, here or public elsewhere, or a local,
    # still binds it.
    module_kinds: dict[str, set[VbaSymbolKind]] = {}
    for symbol in symbols.root.children or ():
        module_kinds.setdefault(symbol.name.lower(), set()).add(symbol.kind)

    def event_only(lower: str) -> bool:
        return lower in module_kinds and all(
            kind is VbaSymbolKind.EVENT for kind in module_kinds[lower]
        )

    def bound(name: str, proc_sym: VbaSymbol | None) -> bool:
        lower = name.lower()
        if not event_only(lower):
            return source_identifier_bound(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.CALL
            )
        return any(
            child.name.lower() == lower for child in (proc_sym.children if proc_sym else None) or ()
        ) or any(
            symbol.kind is not VbaSymbolKind.EVENT and symbol.name.lower() == lower
            for symbol in project_visible_symbols or ()
        )

    def is_known(name: str, proc_sym: VbaSymbol | None) -> bool:
        lower = name.lower()
        return (
            lower in known
            or bound(name, proc_sym)
            or lower in app_members
            or lower in designer_members
            or lower in own_members
            or resolve_host_global(name, host_model) is not None
            # The host's hidden Global interface is bare-callable too (XLIDE #34).
            or resolve_host_global_member(name, host_model) is not None
            or resolve_host_enum(name, host_model) is not None
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
                # A module's name alone is not a call: `Foo` with a module named Foo
                # is "Expected variable or procedure, not module" (XLIDE issue #125,
                # measured in Excel 16.0).
                names_module = any(
                    project_type.kind == "standardModule" and project_type.name.lower() == hit.name.lower()
                    for project_type in project_types or []
                )
                message = (
                    f"'{hit.name}' is a module, not a procedure: name the procedure to call, as in "
                    f"'{hit.name}.Bar'. This is a VBE compile error: Expected variable or procedure, not module."
                    if names_module
                    else f"Sub or Function not defined: '{hit.name}'."
                )
                push("unknownCallStatement", message, hit.span, data)

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
    toks = statement_tokens(source, span)
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
    implicit_members: Sequence[ImplicitMember] | None,
    module_kind: ModuleSymbolKind | None,
    host_model: HostObjectModel | None,
    designer_class: str | None,
    referenced_hosts: Sequence[str] | None,
    push: PushFn,
    own_members: AbstractSet[str] = frozenset(),
) -> None:
    """With Option Explicit, a variable must be declared before it is assigned or
    read. Self-gated on the caller supplying the project-visible identifier set, so
    cross-module globals and enum members never false-positive."""
    if not _has_option_explicit(mod, activity) or known_identifiers is None:
        return
    # A library name or the project name qualifies a global in an expression as it
    # does in an As clause: `Set app = Excel.Application`, `Word.Application`,
    # `VBAProject.Module2.Twice(4)` (XLIDE issue #101; each runs in its host). The
    # libraries are the host's own, those merged into its model (Office, MSForms),
    # the ones the project references, and VBA, which is_known already accepts.
    library_qualifiers = _library_qualifier_names(host_model, referenced_hosts)
    # A form's controls are declared by its DESIGNER, not its text. No control list
    # at all is not an empty one: reading it as empty claimed every control the
    # form's own code-behind names was undeclared (XLIDE issue #48). An Access form
    # or report answers with the list its TypeInfo stream holds, record-source fields
    # included, and the VBE checks a bare name against that list the same way (XLIDE
    # issue #206).
    if module_kind is ModuleSymbolKind.USERFORM and implicit_members is None:
        return
    implicit_member_names = {member.name.lower() for member in implicit_members or ()}

    known = {name.lower() for name in known_identifiers}
    bracket_names_evaluate = _host_evaluates_bracketed_names(host_model)
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    app_members = application_member_names(host_model)
    # The designer's class contributes members the text never declares, and a bare
    # reference to one is correct code.
    designer_members = designer_class_member_names(designer_class, host_model)
    # `ReDim items(2) As Long` at procedure level DECLARES `items` when nothing else
    # does (MS-VBAL 5.4.3.3), and Option Explicit accepts it (XLIDE issue #99, runs
    # in Excel 16.0). Per procedure, below.
    redim_declared: AbstractSet[str] = frozenset()

    def is_known(
        name: str, proc_sym: VbaSymbol | None, context: BareIdentifierContext
    ) -> bool:
        lower = name.lower()
        return (
            lower == "vba"
            or lower in library_qualifiers
            or lower in redim_declared
            # A UserForm's controls are members the designer declared, not the
            # module's text; referring to one is correct VBA.
            or lower in implicit_member_names
            or lower in designer_members
            # The module's own object: UsedRange in a sheet, Tag in a form (XLIDE
            # issue #228).
            or lower in own_members
            or source_identifier_bound(symbols, proc_sym, project_visible_symbols, name, context)
            or lower in known
            or lower in app_members
            or resolve_host_global(name, host_model) is not None
            # Members of the host's hidden Global interface are callable bare
            # (Word's InchesToPoints, Excel's Union), XLIDE #34.
            or resolve_host_global_member(name, host_model) is not None
            or resolve_host_constant(name, host_model) is not None
            # An enum name is a legal qualifier: `XlAxisType.xlCategory` is ordinary
            # VBA, and Option Explicit called the qualifier undeclared.
            or resolve_host_enum(name, host_model) is not None
            or resolve_runtime_constant(name) is not None
            or resolve_runtime_object(name) is not None
            or resolve_runtime_function(name) is not None
            # VBA's own enums and modules qualify their members the same way:
            # `VbMsgBoxResult.vbYes`, `ColorConstants.vbRed`, `Strings.Left`.
            or resolve_vba_library_qualifier(name) is not None
        )

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        redim_declared = _redim_target_names_in(source, member.body, activity)

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
                if ref.bracketed and bracket_names_evaluate:
                    continue
                report(ref.name, ref.span, "using it", BareIdentifierContext.EXPRESSION)

        for_each_undeclared_reference_span(source, member.body, visit, activity)

    # A Const's value and an Enum member's value name things too: `Const K = asdf`
    # is "Variable not defined", and `eB = asdf` in an Enum "Constant expression
    # required" (XLIDE issue #369, measured in Excel 16.0).
    def check_value(
        span: Span, proc_sym: VbaSymbol | None, enum_member: bool, what: str = "a Const's value"
    ) -> None:
        # The names in each value, after its `=`. A value calls nothing, so every
        # name in it not after a `.` is read; a declaration's own names stand
        # before an `=`.
        toks = statement_tokens(source, span)
        in_value = False
        for i, tok in enumerate(toks):
            if tok.raw_text == "=":
                in_value = True
                continue
            if tok.raw_text == ",":
                in_value = False
                continue
            name = token_name(tok) if tok.kind is TokenKind.IDENTIFIER else None
            # A name before a `.` qualifies: `Module2.B1`, `Excel.xlUp`.
            qualifier = _raw_at(toks, i + 1) == "."
            if (
                not in_value
                or not name
                or qualifier
                or _raw_at(toks, i - 1) == "."
                or _raw_at(toks, i - 1) == "!"
                or is_known(name, proc_sym, BareIdentifierContext.EXPRESSION)
            ):
                continue
            push(
                "undeclaredVariable",
                f"'{name}' is not defined, and an Enum member's value must be a constant. "
                "This is a VBE compile error: Constant expression required."
                if enum_member
                else f"Variable not defined: '{name}'. Declare it before using it in {what}, "
                "or remove Option Explicit.",
                Span(span.start + tok.start, span.start + tok.end),
            )

    redim_declared = frozenset()
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and member.is_const:
            check_value(member.span, None, False)
        elif isinstance(member, EnumNode):
            for item in member.members:
                if item.value_raw is not None and not (
                    activity is not None and activity.is_inactive(item.span)
                ):
                    check_value(item.span, None, True)
        elif isinstance(member, ProcedureNode):
            # An Optional parameter's default is a constant from outside the
            # procedure: `Optional x As Long = y` with nothing named y is "Variable
            # not defined" (XLIDE issue #445, measured in Excel 16.0).
            for param in member.params:
                if param.default_raw is not None:
                    check_value(param.span, None, False, "an Optional parameter's default")
            const_proc_sym = procedure_symbol_for(symbols, member)

            def check_const_group(
                group: VariableGroupNode, proc_sym: VbaSymbol | None = const_proc_sym
            ) -> None:
                if group.is_const:
                    check_value(group.span, proc_sym, False)

            for_each_variable_group(member.body, check_const_group, activity)


def _host_evaluates_bracketed_names(host_model: HostObjectModel | None) -> bool:
    """Whether `[name]` on its own is a HOST LOOKUP rather than a variable.

    In Excel the square brackets are shorthand for `Application.Evaluate`, so `[A1]`
    and `[TaxRate]` are ordinary code that compiles and needs no declaration. Word
    has no such feature: `v = [foo]` with nothing declaring `foo` is a compile error
    there, so the report is right for Word and PowerPoint and must stay.

    Only a POSITIVELY identified non-Excel host reports. An absent model is Excel's
    by default, and a host whose model knows nothing asserts nothing, so both of
    those suppress rather than guess.
    """
    host_name = host_model.get("hostName") if host_model is not None else None
    return host_name is None or host_name == "Excel"


def _redim_target_names_in(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> set[str]:
    """Lower-cased names a procedure's ReDim statements size: `ReDim name(...)`,
    `ReDim Preserve name(...)`, and each further `, name(...)`. A ReDim of an
    undeclared name declares it, so these count as declared for the procedure."""
    out: set[str] = set()
    for node in iter_body_nodes(body, inactive_node_skip(activity)):
        if not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            toks = statement_tokens(source, span)
            i = first_executable_token_index(toks)
            if token_text(toks[i] if i < len(toks) else None) != "redim":
                continue
            i += 1
            if token_text(toks[i] if i < len(toks) else None) == "preserve":
                i += 1
            depth = 0
            for k in range(i, len(toks)):
                raw = toks[k].raw_text
                if raw == "(":
                    depth += 1
                elif raw == ")":
                    depth -= 1
                elif depth == 0 and (k == i or toks[k - 1].raw_text == ","):
                    name = token_name(toks[k])
                    if name and k + 1 < len(toks) and toks[k + 1].raw_text == "(":
                        out.add(name.lower())
    return out


def _library_qualifier_names(
    host_model: HostObjectModel | None, referenced_hosts: Sequence[str] | None
) -> set[str]:
    """Lower-cased names that may qualify a global in an expression: the libraries
    whose types the host model carries (its own and the shared ones merged into it),
    the libraries the project references, and the project itself, which is
    `VBAProject` unless renamed. An absent model is Excel's by default."""
    out = {"vbaproject"}
    if host_model is None:
        out.add("excel")
    for qualified in (host_model.get("types") if host_model is not None else None) or {}:
        dot = qualified.find(".")
        if dot > 0:
            out.add(qualified[:dot].lower())
    for token in referenced_hosts or []:
        name = HOST_LIBRARY_NAMES.get(token)
        if name:
            out.add(name.lower())
    return out


_DESIGNER_CLASS_MEMBER_NAMES = IdentityLru(capacity=8)


def designer_class_member_names(
    designer_class: str | None, model: HostObjectModel | None
) -> frozenset[str]:
    """The members of the class a module's designer makes it, lowercased. Inside such
    a module they are in scope unqualified, exactly as the module's own procedures
    are, because the module IS one of these. Empty when the caller cannot say which
    class it is, or the model does not carry that type."""
    if not designer_class or model is None:
        return frozenset()
    by_type: dict[str, frozenset[str]] | None = _DESIGNER_CLASS_MEMBER_NAMES.get(model)
    if by_type is None:
        by_type = {}
        _DESIGNER_CLASS_MEMBER_NAMES.put(by_type, model)
    key = designer_class.lower()
    names = by_type.get(key)
    if names is None:
        names = frozenset(member["name"].lower() for member in get_host_members(designer_class, model))
        by_type[key] = names
    return names


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
