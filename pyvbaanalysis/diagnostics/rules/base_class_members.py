"""Rule: member-exists-in-base-class.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/baseClassMembers.ts.

A document module, an Access form or report module, and a UserForm's module each
derive from a host class: ThisWorkbook from Workbook, a sheet from Worksheet or
Chart, ThisDocument from Document, Form_X from Form, Report_X from Report, a
UserForm from UserForm. A module-level declaration may not take a name that class
already has: the VBE refuses it with "Member already exists in an object module
from which this object module derives". Measured in Excel, Word, PowerPoint and
Access 16.0 (2026-10-16):

  - It clashes: a Sub, Function or Property Get, Let or Set, Public, Private,
    Friend or with no modifier; a module-level variable, Dim, Private or Public;
    a Private Const; a Private Declare. Any letter case, and hidden members count
    (Document.Dummy1).
  - It does not: a WithEvents variable, an Enum or its members, a Type, an Event,
    a local, a Static local, a parameter, a label, or a name only the Application
    has (Selection, CurrentDb).

Not reported here, because the VBE refuses them for another reason first: a
keyword such as Close ("Expected: identifier"), and a Public Const, Declare or
array, which an object module may not make public at all.

A sheet module whose sheet type is unknown is held only to the members a Worksheet
and a Chart share. VB6's designer classes were not measured, so a VB6 module is
left alone.
"""

from __future__ import annotations

import re
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...host.host_model import get_excel_object_model
from ...identity_cache import IdentityLru
from ...lexer.keyword_table import is_reserved_identifier
from ...parser.nodes import DeclareNode, ModuleNode, ProcedureNode, Span, VariableGroupNode
from ...symbols.symbol_model import ModuleSymbolKind
from ..context import AnalyzeModuleOptions, PushFn, own_object_member_names
from ..walker import active_module_members
from .undeclared import designer_class_member_names

# The Office designer classes this rule was measured against.
_MEASURED_DESIGNER_CLASS = re.compile(r"^(Excel|Word|PowerPoint|Access)\.", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class _BaseClass:
    # What the message calls the class: `Word.Document`, `UserForm`.
    label: str
    # Its members, lowercased.
    members: AbstractSet[str]


# What a Worksheet and a Chart share, per model: built once, read per analysis.
_SHEET_MEMBERS = IdentityLru(capacity=8)


def _base_class_of(opts: AnalyzeModuleOptions) -> _BaseClass | None:
    """The host class the module derives from, when it is known for certain."""
    if opts.designer_class:
        if not _MEASURED_DESIGNER_CLASS.search(opts.designer_class):
            return None
        model = opts.host_model if opts.host_model is not None else get_excel_object_model()
        return _BaseClass(opts.designer_class, designer_class_member_names(opts.designer_class, model))
    if opts.module_kind is ModuleSymbolKind.USERFORM:
        if (opts.host or "").lower() == "vb6":
            return None
        return _BaseClass("UserForm", own_object_member_names(opts))
    if opts.module_kind is not ModuleSymbolKind.DOCUMENT:
        return None
    model = opts.host_model if opts.host_model is not None else get_excel_object_model()
    host = (model.get("hostName") or "Excel").lower()
    name = opts.module_name.lower() if opts.module_name is not None else None
    document_type = opts.document_type
    known: str | None
    if document_type == "worksheet":
        known = "Excel.Worksheet"
    elif document_type == "chart":
        known = "Excel.Chart"
    elif document_type == "workbook":
        known = "Excel.Workbook"
    elif document_type == "document":
        known = "Word.Document"
    elif host == "excel" and name == "thisworkbook":
        known = "Excel.Workbook"
    elif host == "word" and name == "thisdocument":
        known = "Word.Document"
    else:
        known = None
    if known:
        return _BaseClass(known, designer_class_member_names(known, model))
    if host != "excel":
        return None
    # A sheet of either kind: only what both kinds have is certain to clash.
    shared: frozenset[str] | None = _SHEET_MEMBERS.get(model)
    if shared is None:
        chart = designer_class_member_names("Excel.Chart", model)
        shared = frozenset(
            member for member in designer_class_member_names("Excel.Worksheet", model) if member in chart
        )
        _SHEET_MEMBERS.put(shared, model)
    return _BaseClass("Excel.Worksheet", shared)


def check_base_class_member_clash(
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    opts: AnalyzeModuleOptions,
    push: PushFn,
) -> None:
    base = _base_class_of(opts)
    if base is None or not base.members:
        return

    def report(name: str, span: Span | None) -> None:
        if span is None or is_reserved_identifier(name) or name.lower() not in base.members:
            return
        push(
            "memberExistsInBaseClass",
            f"'{name}' is already a member of {base.label}, which this module derives from. "
            "This is a VBE compile error: Member already exists in an object module from which "
            "this object module derives.",
            span,
        )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            report(member.name, member.name_span)
        elif isinstance(member, DeclareNode):
            if (member.visibility or "").lower() != "public":
                report(member.name, member.name_span)
        elif isinstance(member, VariableGroupNode) and not member.with_events:
            is_public = member.modifier.lower() in ("public", "global")
            for decl in member.declarations:
                if is_public and (member.is_const or decl.is_array or decl.fixed_length is not None):
                    continue
                report(decl.name, decl.name_span)
