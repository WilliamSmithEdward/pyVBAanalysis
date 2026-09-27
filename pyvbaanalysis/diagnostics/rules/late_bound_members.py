"""Rule: a member the VBE binds at run time, on an object whose class the code
makes plain and whose member list is complete (XLIDE issue #121).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lateBoundMembers.ts.
Measured in Excel 16.0 (build 20326, 2026-09-26); each compiles and raises 438,
"Object doesn't support this property or method", every time it runs.

 - `Application.Zzq`: Application is extensible, so the VBE compiles any
   name on it (worksheet functions such as Application.Match are ordinary
   VBA there), and a name that is neither an Application member nor a
   WorksheetFunction raises when it runs. Excel only: its model lists every
   member, hidden ones included.
 - `Dim o As Object: Set o = New Collection: o.Foo`: a late-bound variable
   holding a class with a known member list. Collection has Add, Count,
   Item and Remove; a project class module has its public members.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext, resolve_receiver_type_at
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declaration
from ...host.host_model import HostObjectModel, get_host_members, get_host_type
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import (
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..context import PushFn
from ..walker import (
    active_module_members,
    set_assignment_target,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_COLLECTION_MEMBERS: frozenset[str] = frozenset({"add", "count", "item", "remove"})


@dataclass(frozen=True, slots=True)
class _KnownClass:
    """Names a late-bound local is known to hold: the class display name and its members."""

    display: str
    members: AbstractSet[str]


def check_runtime_member_not_found(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`Application.Zzq` in Excel, and `o.Foo` on an Object or Variant local that
    holds a New Collection or an exhaustive project class, raise 438 at run time."""
    model = member_ctx.model
    application_surface = _excel_application_surface(model)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        # Asked only for the target of a Set: walking the whole environment for
        # every procedure was 5% of a large module's pass (XLIDE issue #139).
        def is_late_bound(lower: str, env: Mapping[str, str] = env) -> bool:
            if lower not in env:
                return False
            normalized = normalize_type(env[lower])
            return normalized == "object" or normalized == "variant" or normalized is None

        held: dict[str, _KnownClass] = {}
        for node in member.body:
            if activity is not None and activity.is_inactive(node.span):
                continue
            if isinstance(node, VariableGroupNode):
                continue  # a Dim inside the body declares, and runs nothing
            if not is_leaf_statement(node):
                held.clear()
                continue
            toks = statement_tokens_after_leading_label(source, node.span)
            if (
                statement_label_declaration(source, node.span) is not None
                or token_text(toks[0] if toks else None) == "gosub"
            ):
                held.clear()
            if isinstance(node, StatementNode) and node.single_line_if_branches is not None:
                _forget_mentioned(toks, held)
                continue
            _check_statement(
                source, node.span.start, toks, held, application_surface, member_ctx, push
            )
            assigned = set_assignment_target(source, node.span)
            if assigned is not None and is_late_bound(assigned[0].lower()):
                lower = assigned[0].lower()
                value = toks[_index_of_equals(toks) + 1 :]
                known = (
                    _known_class_named(token_name(value[1]), member_ctx)
                    if len(value) == 2 and token_text(value[0]) == "new"
                    else None
                )
                if known is not None:
                    held[lower] = known
                else:
                    held.pop(lower, None)
                continue
            _forget_other_uses(toks, held)


def _known_class_named(
    name: str | None, member_ctx: MemberCompletionContext
) -> _KnownClass | None:
    if not name:
        return None
    if name.lower() == "collection":
        return _KnownClass("Collection", _COLLECTION_MEMBERS)
    project_type = next(
        (
            candidate
            for candidate in member_ctx.project_class_members or []
            if candidate.kind == "class"
            and candidate.exhaustive is True
            and candidate.name.lower() == name.lower()
        ),
        None,
    )
    if project_type is None:
        return None
    return _KnownClass(
        project_type.name, frozenset(m.name.lower() for m in project_type.members)
    )


def _excel_application_surface(model: HostObjectModel | None) -> AbstractSet[str] | None:
    """Excel's Application members plus the worksheet functions it also answers to."""
    if model is not None and "hostName" in model and model["hostName"] != "Excel":
        return None
    application = get_host_type("Excel.Application", model)
    if application is None or application.get("exhaustive") is not True:
        return None
    names: set[str] = set()
    for member in get_host_members("Excel.Application", model):
        names.add(member["name"].lower())
    for member in get_host_members("Excel.WorksheetFunction", model):
        names.add(member["name"].lower())
    return names


def _check_statement(
    source: str,
    base: int,
    toks: Sequence[VbaToken],
    held: dict[str, _KnownClass],
    application_surface: AbstractSet[str] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    for i in range(len(toks) - 2):
        if toks[i + 1].raw_text != "." or (i >= 1 and toks[i - 1].raw_text == "."):
            continue
        receiver = token_name(toks[i])
        member_name = token_name(toks[i + 2])
        if not receiver or not member_name:
            continue
        at = Span(base + toks[i + 2].start, base + toks[i + 2].end)
        known = held.get(receiver.lower())
        if known is not None:
            if member_name.lower() not in known.members:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, which has no member "
                    f"'{member_name}'. This will raise Run-time error '438': Object doesn't "
                    "support this property or method.",
                    at,
                )
            continue
        if (
            application_surface is not None
            and receiver.lower() == "application"
            and member_name.lower() not in application_surface
        ):
            receiver_type = resolve_receiver_type_at(source, base + toks[i + 1].end, member_ctx)
            if receiver_type == "Excel.Application":
                push(
                    "runtimeMemberNotFound",
                    f"Application has no member '{member_name}', and it is not a worksheet "
                    "function either. The VBE compiles the name because Application is "
                    "extensible; this will raise Run-time error '438': Object doesn't "
                    "support this property or method.",
                    at,
                )


def _forget_other_uses(toks: Sequence[VbaToken], held: dict[str, _KnownClass]) -> None:
    """A tracked variable named in any position other than `name.Member` is no longer followed."""
    for i in range(len(toks)):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        if (
            lower
            and lower in held
            and not (i >= 1 and toks[i - 1].raw_text == ".")
            and not (i + 1 < len(toks) and toks[i + 1].raw_text == ".")
        ):
            del held[lower]


def _forget_mentioned(toks: Sequence[VbaToken], held: dict[str, _KnownClass]) -> None:
    for tok in toks:
        name = token_name(tok)
        lower = name.lower() if name is not None else None
        if lower and lower in held:
            del held[lower]


def _index_of_equals(toks: Sequence[VbaToken]) -> int:
    """Index of the first `=` token, or -1."""
    for k, tok in enumerate(toks):
        if tok.raw_text == "=":
            return k
    return -1
