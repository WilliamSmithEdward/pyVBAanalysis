"""Rule family: duplicate and ambiguous declarations.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/duplicates.ts. Covers the
five duplicate-declaration rules plus the cross-module ambiguous-enum-member rule
(checkAmbiguousEnumMemberReferences): an unqualified read of a member name shared
by more than one visible Enum is the VBA "Ambiguous name detected" compile error.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...conditional import ConditionalActivityTracker
from ..call_extraction import extract_call
from ..walker import ProcedureStatementVisitor
from ..callable_signatures import (
    bare_callable_source_shadowed,
    same_module_callable_signatures,
    source_name_scope_for,
)
from ...host import application_member_names, resolve_host_global
from ...host.host_model import HostObjectModel
from ...parser.nodes import (
    EnumMemberNode,
    EnumNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    TypeFieldNode,
    TypeNode,
)
from ...runtime import resolve_runtime_function, resolve_runtime_object
from ...symbols.name_resolution import (
    BareIdentifierContext,
    BareIdentifierResolution,
    BareIdentifierResolutionScope,
)
from ...symbols.symbol_model import (
    ModuleSymbols,
    SymbolVisibility,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
    is_procedure_kind,
)
from ...types.type_inference import (
    procedure_symbol_for,
    source_identifier_binding,
    source_identifier_bound,
)
from ..callable_signatures import callable_type_signatures_for
from ..context import PushFn
from ..walker import active_module_members
from .shared import (
    declaration_name_hit,
    for_each_undeclared_reference_span,
    report_repeated_keys,
    value_read_references,
)

_PROPERTY_KINDS = (
    VbaSymbolKind.PROPERTY_GET,
    VbaSymbolKind.PROPERTY_LET,
    VbaSymbolKind.PROPERTY_SET,
)
_LOCAL_DECL_KINDS = (
    VbaSymbolKind.PARAMETER,
    VbaSymbolKind.LOCAL_VARIABLE,
    VbaSymbolKind.CONSTANT,
)


def _ALWAYS_COLLIDE(a: VbaSymbol, b: VbaSymbol) -> bool:
    """Two declarations of one name always collide; only the scope differs."""
    return True


def _report_repeated_names(
    symbols: Sequence[VbaSymbol],
    activity: ConditionalActivityTracker | None,
    declares: Callable[[VbaSymbol], bool],
    collides: Callable[[VbaSymbol, VbaSymbol], bool],
    report: Callable[[VbaSymbol], None],
) -> None:
    """Report each symbol whose name an EARLIER symbol already took.

    A repeat only counts when the two could be compiled together: the arms of one
    `#If` chain are alternatives, not duplicate declarations, since no build ever
    sees both (XLIDE issue #58). Almost every name is taken once, so this keeps
    one list per name and only grows it when a second symbol claims it.
    """
    taken: dict[str, list[VbaSymbol]] = {}
    for sym in symbols:
        if not declares(sym):
            continue
        key = sym.name.lower()
        earlier = taken.get(key)
        if earlier is None:
            taken[key] = [sym]
            continue
        hit = any(
            collides(prior, sym)
            and (activity is None or not activity.mutually_exclusive(prior.name_span, sym.name_span))
            for prior in earlier
        )
        earlier.append(sym)
        if hit:
            report(sym)


def _procedures_collide(a: VbaSymbol, b: VbaSymbol) -> bool:
    """Distinct accessors of one property share their name legitimately; every
    other repeat is the ambiguity error."""
    return a.kind not in _PROPERTY_KINDS or b.kind not in _PROPERTY_KINDS or a.kind is b.kind


def check_duplicate_procedures(
    members: Sequence[VbaSymbol], activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """A name may be one Sub/Function OR a set of distinct Property accessors."""
    def report(sym: VbaSymbol) -> None:
        push(
            "duplicateProcedure",
            f"Ambiguous name detected: '{sym.name}' is already declared in this module.",
            sym.name_span,
        )

    _report_repeated_names(
        members, activity, lambda sym: is_procedure_kind(sym.kind), _procedures_collide, report
    )


def check_duplicate_declarations(
    members: Sequence[VbaSymbol], activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Within one procedure, a name is declared once across params/locals/consts.

    Procedure scope is flat in VBA (no block scope), so locals from different `If`
    branches still collide, but locals from different `#If` arms do not, because
    only one of those arms is ever compiled."""
    for proc in members:
        if not is_procedure_kind(proc.kind):
            continue
        def report(sym: VbaSymbol) -> None:
            push(
                "duplicateDeclaration",
                f"Duplicate declaration in current scope: '{sym.name}'.",
                sym.name_span,
            )

        _report_repeated_names(
            proc.children or [], activity, lambda sym: sym.kind in _LOCAL_DECL_KINDS,
            _ALWAYS_COLLIDE, report,
        )


def check_duplicate_module_members(
    members: Sequence[VbaSymbol], activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """A module-level variable or constant declared more than once."""
    def report(sym: VbaSymbol) -> None:
        push(
            "duplicateModuleMember",
            f"Duplicate declaration: '{sym.name}' is already declared at module level.",
            sym.name_span,
        )

    _report_repeated_names(
        members, activity,
        lambda sym: sym.kind in (VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.CONSTANT),
        _ALWAYS_COLLIDE, report,
    )


def _block_entry_key(
    entry: EnumMemberNode | TypeFieldNode, activity: ConditionalActivityTracker | None
) -> str | None:
    """The repeat key of an Enum member or Type field. Only a provably inactive one
    drops out: an entry in a branch that cannot be decided still collides with
    anything the same build would compile beside it, and only the arms of one
    `#If` chain are alternatives. Skipping every undecidable branch went blind to a
    genuine repeat inside one arm."""
    return None if activity is not None and activity.is_inactive(entry.span) else entry.name.lower()


def check_duplicate_enum_members(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Member names inside one Enum block must be unique."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, EnumNode):
            continue

        def report(repeat: EnumMemberNode, earlier: EnumMemberNode, enum_name: str = member.name) -> None:
            hit = declaration_name_hit(source, repeat.span, repeat.name)
            push(
                "duplicateEnumMember",
                f"Duplicate Enum member '{repeat.name}' in Enum '{enum_name}'.",
                hit.span if hit is not None else repeat.span,
            )

        report_repeated_keys(
            member.members,
            activity,
            lambda entry: _block_entry_key(entry, activity),
            lambda entry: entry.span,
            report,
        )


def check_duplicate_type_fields(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """Field names inside one Type (UDT) block must be unique."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, TypeNode):
            continue

        def report(repeat: TypeFieldNode, earlier: TypeFieldNode, type_name: str = member.name) -> None:
            hit = declaration_name_hit(source, repeat.span, repeat.name)
            push(
                "duplicateTypeField",
                f"Duplicate field '{repeat.name}' in Type '{type_name}'.",
                hit.span if hit is not None else repeat.span,
            )

        report_repeated_keys(
            member.fields,
            activity,
            lambda entry: _block_entry_key(entry, activity),
            lambda entry: entry.span,
            report,
        )


def check_ambiguous_enum_member_references(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    module_name: str,
    known_procedures: AbstractSet[str] | None,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_members: Sequence[VbaProjectClassMembers] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    host_model: HostObjectModel | None,
    push: PushFn,
) -> None:
    """An unqualified read of a member name shared by more than one visible Enum is
    rejected as "Ambiguous name detected".

    Same-module bindings take precedence over exported members from other modules,
    and procedure locals/parameters shadow module-level enum members. The
    no-false-positive hinge is the binder: a read that binds to a local/param or any
    non-enumMember symbol resolves with scope != AMBIGUOUS and stays silent; only an
    all-enumMember binding owned by more than one distinct module:container fires.
    """
    visible_enum_members = [
        *_enum_member_symbols(symbols.root.children or []),
        *[
            sym
            for sym in (project_visible_symbols or [])
            if sym.kind is VbaSymbolKind.ENUM_MEMBER
            and sym.module_name.lower() != module_name.lower()
        ],
    ]
    if len(_ambiguous_enum_member_groups(visible_enum_members)) == 0:
        return

    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    app_members = application_member_names(host_model)
    known = {name.lower() for name in (known_procedures or ())}

    def is_known_for_skip(name: str, proc_sym: VbaSymbol | None) -> bool:
        lower = name.lower()
        return (
            source_identifier_bound(
                symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.EXPRESSION
            )
            or lower in known
            or lower in app_members
            or resolve_host_global(name, host_model) is not None
            or resolve_runtime_object(name) is not None
            or resolve_runtime_function(name) is not None
        )

    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        reported: set[str] = set()

        def visit(
            span: Span,
            proc_sym: VbaSymbol | None = proc_sym,
            reported: set[str] = reported,
        ) -> None:
            def is_skipped(name: str) -> bool:
                return is_known_for_skip(name, proc_sym)

            for ref in value_read_references(
                source,
                span,
                is_skipped,
                module_signatures,
                project_members,
            ):
                binding = source_identifier_binding(
                    symbols,
                    proc_sym,
                    project_visible_symbols,
                    ref.name,
                    BareIdentifierContext.EXPRESSION,
                )
                definitions = _ambiguous_enum_member_definitions(binding)
                if definitions is None:
                    continue
                key = f"{ref.span.start}:{ref.span.end}"
                if key in reported:
                    continue
                reported.add(key)
                owners: list[str] = []
                for definition in definitions:
                    owner = definition.container_name or definition.module_name
                    if owner not in owners:
                        owners.append(owner)
                owner_text = ", ".join(owners[:3])
                detail = f" ({owner_text})" if owner_text else ""
                push(
                    "ambiguousEnumMember",
                    f"Ambiguous Enum member reference: '{ref.name}' is defined by "
                    f"multiple visible Enums{detail}. Qualify the reference with an "
                    f"Enum or module name.",
                    ref.span,
                )

        for_each_undeclared_reference_span(source, member.body, visit, activity)


def _enum_member_symbols(symbols: Sequence[VbaSymbol]) -> list[VbaSymbol]:
    out: list[VbaSymbol] = []
    for symbol in symbols:
        if symbol.kind is VbaSymbolKind.ENUM:
            out.extend(
                child
                for child in (symbol.children or [])
                if child.kind is VbaSymbolKind.ENUM_MEMBER
            )
    return out


def _ambiguous_enum_member_groups(
    symbols: Sequence[VbaSymbol],
) -> dict[str, list[VbaSymbol]]:
    groups: dict[str, dict[str, VbaSymbol]] = {}
    for symbol in symbols:
        if symbol.kind is not VbaSymbolKind.ENUM_MEMBER:
            continue
        key = symbol.name.lower()
        owner_key = f"{symbol.module_name.lower()}:{(symbol.container_name or '').lower()}"
        owners = groups.setdefault(key, {})
        owners.setdefault(owner_key, symbol)
    return {key: list(owners.values()) for key, owners in groups.items() if len(owners) > 1}


def _ambiguous_enum_member_definitions(
    binding: BareIdentifierResolution,
) -> list[VbaSymbol] | None:
    if binding.scope is not BareIdentifierResolutionScope.AMBIGUOUS:
        return None
    if any(d.kind is not VbaSymbolKind.ENUM_MEMBER for d in binding.definitions):
        return None
    owner_keys = {
        f"{d.module_name.lower()}:{(d.container_name or '').lower()}" for d in binding.definitions
    }
    return list(binding.definitions) if len(owner_keys) > 1 else None


# -- checkAmbiguousBareProcedureCalls --------------------------------------


def _ambiguous_project_procedure_owners(
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    module_name: str,
) -> dict[str, list[str]]:
    """Names exported by more than one OTHER module, mapped to those module names.

    The calling module is excluded because its own declaration would settle the
    name before the project is consulted.
    """
    out: dict[str, list[str]] = {}
    if not project_procedures:
        return out
    self_name = module_name.lower()
    for name, signatures in project_procedures.items():
        owners: list[str] = []
        for signature in signatures:
            # A Private procedure is not exported, so it cannot collide.
            if signature.visibility is SymbolVisibility.PRIVATE:
                continue
            if not any(owner.lower() == signature.module_name.lower() for owner in owners):
                owners.append(signature.module_name)
        if len(owners) > 1 and not any(owner.lower() == self_name for owner in owners):
            out[name.lower()] = owners
    return out


def check_ambiguous_bare_procedure_calls(
    source: str,
    symbols: ModuleSymbols,
    module_name: str,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """VBA is content for two modules to export the same public procedure name, but
    it refuses to compile an UNQUALIFIED call to that name from a module declaring
    neither: "Ambiguous name detected".

    The finding belongs at the CALL SITE, not the declarations: a project that
    exports a name twice and always qualifies its calls is legal VBA and common, so
    flagging the declarations would cry wolf on every one of them.

    Silent, matching VBA, when any of these settle the name: the call is qualified
    (`Helpers.Recalculate`); the calling module declares the name itself (module-local
    scope wins); a local, parameter or module-level symbol shadows it; or only one
    module in the project exports it.
    """
    ambiguous_names = _ambiguous_project_procedure_owners(project_procedures, module_name)
    same_module_signatures = same_module_callable_signatures(symbols)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        # No name in this project is exported twice: nothing here can be
        # ambiguous, so skip the per-statement work entirely.
        if not ambiguous_names:
            return None
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            call = extract_call(source, stmt.span)
            if call is None or call.qualifier:
                return
            lower = call.name.lower()
            owners = ambiguous_names.get(lower)
            if owners is None:
                return
            # This module declares it, so VBA binds locally and never asks.
            if lower in same_module_signatures:
                return
            if bare_callable_source_shadowed(call.name, source_names):
                return
            push(
                "ambiguousProjectProcedure",
                f"Ambiguous name detected: '{call.name}' is exported by "
                f"{' and '.join(owners)}. VBA refuses to compile the project until this "
                "call is qualified with a module name.",
                call.name_span,
            )

        return visitor

    return factory
