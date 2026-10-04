"""Rule: an event handler's declaration must match its event (XLIDE issue #195).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/eventHandlerSignatures.ts.
Measured in Excel 16.0 (build 20326, 2026-09-29): a handler that differs from
its event is refused with "Procedure declaration does not match description of
event or procedure having the same name".

- Each parameter's passing: `app_SheetChange(Sh As Object, ...)` is
  refused, since the event passes Sh ByVal. An event's ByRef parameter,
  `Cancel As Boolean`, may be written ByRef or plain, not ByVal.
- Each parameter's type: Variant, a missing As, or Worksheet for the
  event's Object is refused. A library prefix is the same type
  (`Excel.Range`), and an enum parameter may be declared As Long.
- The count, and no Optional or ParamArray in place of the event's own.
- A Sub: a Function of the event's name is refused.

Names differ freely, and Public or Private both compile. The VBE checks a
handler only when its body holds something: an empty one compiles
whatever its declaration.

The handlers judged: `<variable>_<Event>` for a WithEvents variable of an
Office or MSForms class, a document module's own object (Workbook_,
Worksheet_, Chart_, Word's Document_), a UserForm's (UserForm_) and its
controls'. The events come from the type libraries
(host/event_signatures_data.py).

A WithEvents variable of a project class takes the events the class
declares (XLIDE issue #220, measured in Excel 16.0). The same checks hold, an
empty handler compiles here too, and a parameter of a project Enum may be
declared As Long.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from dataclasses import dataclass

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...host.event_signatures_data import HOST_EVENT_SIGNATURES
from ...host.host_model import get_host_type, resolve_host_alias
from ...js_compat import JS_WHITESPACE, js_trim
from ...parser.nodes import ModuleNode, ParameterNode, ProcedureNode, ProcKind, VariableGroupNode
from ...symbols.symbol_model import ModuleSymbolKind, is_access_designer_class
from ...types.type_names import normalize_type
from ..context import AnalyzeModuleOptions, PushFn
from ..walker import active_module_members


@dataclass(frozen=True, slots=True)
class _Event:
    name: str
    params: str


@dataclass(frozen=True, slots=True)
class _EventClass:
    name: str
    events: Mapping[str, _Event]


def _built_in_events() -> dict[str, _EventClass]:
    """The events of each class, by lowercased qualified name, read once."""
    out: dict[str, _EventClass] = {}
    for class_name, events in HOST_EVENT_SIGNATURES.items():
        by_name = {event.lower(): _Event(event, params) for event, params in events.items()}
        out[class_name.lower()] = _EventClass(class_name, by_name)
    return out


_EVENTS_BY_CLASS: dict[str, _EventClass] = _built_in_events()

# The events VBA's own UserForm adds to FM20's, which the type library sweep
# cannot see: a handler with parameters, or a QueryClose passing either one
# ByVal or as another type, is refused (XLIDE issue #226, measured in Excel 16.0).
_VBA_USERFORM = "VBA.UserForm"
_EVENTS_BY_CLASS[_VBA_USERFORM.lower()] = _EventClass(
    _VBA_USERFORM,
    {
        name.lower(): _Event(name, params)
        for name, params in {
            "Initialize": "",
            "Terminate": "",
            "Activate": "",
            "Deactivate": "",
            "Resize": "",
            "QueryClose": "Cancel As Integer, CloseMode As Integer",
        }.items()
    },
)

# A document module's own object: the prefix its handlers take, and its class.
_DOCUMENT_OBJECTS: Mapping[str, tuple[str, str]] = {
    "workbook": ("Workbook", "Excel.Workbook"),
    "worksheet": ("Worksheet", "Excel.Worksheet"),
    "chart": ("Chart", "Excel.Chart"),
    "document": ("Document", "Word.Document"),
}

_HOST_ENUM_RE = re.compile(r"^(xl|wd|pp|mso|fm)[a-z]", re.IGNORECASE | re.ASCII)


def check_event_handler_signatures(
    mod: ModuleNode,
    module_kind: ModuleSymbolKind,
    opts: AnalyzeModuleOptions,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    # Each prefix a handler may take, with the classes whose events it raises.
    sources: dict[str, list[str]] = {}
    project_sources: dict[str, tuple[str, dict[str, _Event]]] = {}

    def add(prefix: str, *classes: str) -> None:
        known = [class_name for class_name in classes if class_name.lower() in _EVENTS_BY_CLASS]
        if len(known) > 0:
            sources[prefix.lower()] = known

    # A document module's own object. Its handlers are checked even when
    # empty, unlike a WithEvents variable's, a project class's or a form's
    # (XLIDE issue #228, measured in Excel 16.0).
    document_prefixes: set[str] = set()
    module_lower = opts.module_name.lower() if opts.module_name is not None else None
    document_type = (
        opts.document_type
        if opts.document_type is not None
        else "workbook"
        if module_lower == "thisworkbook"
        else "document"
        if module_lower == "thisdocument"
        else None
    )
    document_object = (
        _DOCUMENT_OBJECTS.get(document_type) if isinstance(document_type, str) else None
    )
    if module_kind is ModuleSymbolKind.DOCUMENT and document_type and document_object:
        add(*document_object)
        document_prefixes.add(document_object[0].lower())
    if module_kind is ModuleSymbolKind.USERFORM and is_access_designer_class(opts.designer_class):
        # An Access form or report, and its controls, raise Access's events:
        # `Qty_BeforeUpdate(Cancel As Integer)` compiles there, and
        # `Form_BeforeUpdate(ByVal Cancel As Integer)` does not (XLIDE issue #227).
        designer = opts.designer_class or ""
        _add_host_events(
            "Report" if designer.lower() == "access.report" else "Form",
            designer,
            member_ctx,
            sources,
        )
        for control in opts.implicit_members or []:
            _add_host_events(control.name, control.type, member_ctx, sources)
    elif module_kind is ModuleSymbolKind.USERFORM:
        add("UserForm", "MSForms.UserForm", _VBA_USERFORM)
        # A control's handlers take its own events and the extender's
        # (Enter, Exit, BeforeUpdate, AfterUpdate).
        for control in opts.implicit_members or []:
            add(control.name, control.type, "MSForms.Control")
    procedures: list[ProcedureNode] = []
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and member.with_events:
            for decl in member.declarations:
                project = _project_events_of(decl.as_type, member_ctx)
                if project is not None:
                    project_sources[decl.name.lower()] = project
                    continue
                class_name = _event_class_of(decl.as_type, member_ctx)
                if class_name:
                    add(decl.name, class_name)
        elif isinstance(member, ProcedureNode):
            procedures.append(member)
    if len(sources) == 0 and len(project_sources) == 0:
        return
    project_enums = {
        type_.name.lower()
        for type_ in (member_ctx.project_class_members or [])
        if type_.kind == "enum"
    }
    for proc in procedures:
        underscore = proc.name.rfind("_")
        if underscore <= 0:
            continue
        prefix = proc.name[:underscore].lower()
        # The VBE checks a handler only when its body holds a statement, a
        # document's own handlers excepted.
        if len(proc.body) == 0 and prefix not in document_prefixes:
            continue
        event_name = proc.name[underscore + 1 :].lower()
        project_source = project_sources.get(prefix)
        project_event = project_source[1].get(event_name) if project_source is not None else None
        if project_source is not None and project_event is not None:
            problem = _mismatch(
                proc, project_event.params, lambda type_: _bare_type(type_) in project_enums
            )
            if problem:
                push(
                    "eventHandlerSignature",
                    f"'{proc.name}' does not match the event {project_source[0]}."
                    f"{project_event.name}({project_event.params}): {problem}. This is a VBE "
                    "compile error: Procedure declaration does not match description of event or "
                    "procedure having the same name.",
                    proc.name_span if proc.name_span is not None else proc.span,
                )
            continue
        classes = sources.get(prefix)
        found: tuple[_EventClass, _Event] | None = None
        for class_name in classes or []:
            owner = _EVENTS_BY_CLASS[class_name.lower()]
            event = owner.events.get(event_name)
            if event is not None:
                found = (owner, event)
                break
        if found is None:
            continue
        owner, event = found
        problem = _mismatch(proc, event.params, lambda type_: bool(_HOST_ENUM_RE.search(type_)))
        if problem:
            owner_name = owner.name[owner.name.find(".") + 1 :]
            push(
                "eventHandlerSignature",
                f"'{proc.name}' does not match the event {owner_name}.{event.name}({event.params}): "
                f"{problem}. This is a VBE compile error: Procedure declaration does not match "
                "description of event or procedure having the same name.",
                proc.name_span if proc.name_span is not None else proc.span,
            )


def _project_events_of(
    as_type: str | None, member_ctx: MemberCompletionContext
) -> tuple[str, dict[str, _Event]] | None:
    """The events a project class declares, when a WithEvents variable's type
    names one: the project's types come before the libraries', as in VBA."""
    trimmed = js_trim(as_type) if as_type is not None else None
    bare = trimmed[trimmed.rfind(".") + 1 :].lower() if trimmed is not None else None
    type_ = (
        next(
            (
                candidate
                for candidate in (member_ctx.project_class_members or [])
                if candidate.name.lower() == bare
            ),
            None,
        )
        if bare
        else None
    )
    if type_ is None:
        return None
    events: dict[str, _Event] = {}
    for member in type_.members:
        signature = member.signature if member.signature is not None else ""
        open_ = signature.find("(")
        if member.kind == "event" and open_ >= 0 and signature.endswith(")"):
            events[member.name.lower()] = _Event(member.name, signature[open_ + 1 : -1])
    return type_.name, events


def _add_host_events(
    prefix: str, type_: str, member_ctx: MemberCompletionContext, sources: dict[str, list[str]]
) -> None:
    """The events a host type declares in its object model, under the prefix a
    handler takes: Access keeps its events there, read from MSACC.OLB with
    ByVal where the library passes by value."""
    alias = resolve_host_alias(type_, member_ctx.model)
    resolved = alias if alias is not None else type_
    key = resolved.lower()
    if key not in _EVENTS_BY_CLASS:
        events: dict[str, _Event] = {}
        # The type's own list: the member index leaves events out.
        host_type = get_host_type(resolved, member_ctx.model)
        for member in (host_type.get("members") if host_type is not None else None) or []:
            signature = member.get("signature")
            open_ = signature.find("(") if signature is not None else -1
            if (
                member.get("kind") == "event"
                and signature
                and open_ >= 0
                and signature.endswith(")")
            ):
                events[member["name"].lower()] = _Event(member["name"], signature[open_ + 1 : -1])
        if len(events) == 0:
            return
        # Upstream memoizes these into its module-level class table, as here:
        # each entry depends only on the resolved type's name.
        _EVENTS_BY_CLASS[key] = _EventClass(resolved, events)
    sources[prefix.lower()] = [resolved]


def _event_class_of(as_type: str | None, member_ctx: MemberCompletionContext) -> str | None:
    """The class with events a WithEvents variable's declared type names, or None."""
    if not as_type:
        return None
    direct = _EVENTS_BY_CLASS.get(as_type.lower())
    if direct is not None:
        return direct.name
    resolved = resolve_host_alias(as_type, member_ctx.model)
    return resolved if resolved and resolved.lower() in _EVENTS_BY_CLASS else None


@dataclass(frozen=True, slots=True)
class _EventParam:
    by_val: bool
    type: str
    is_array: bool


_WS = JS_WHITESPACE
_AS_TYPE_RE = re.compile(f"[{_WS}]As[{_WS}]+([^{_WS}]+)\\Z", re.IGNORECASE | re.ASCII)
_PASSING_PREFIX_RE = re.compile(f"^(?:ByVal|ByRef)[{_WS}]+", re.IGNORECASE | re.ASCII)
_WHITESPACE_CHAR_RE = re.compile(f"[{_WS}]")
_BYVAL_RE = re.compile(f"^ByVal[{_WS}]", re.IGNORECASE | re.ASCII)


def _parse_event_params(params: str) -> list[_EventParam]:
    if len(js_trim(params)) == 0:
        return []
    out: list[_EventParam] = []
    for part in params.split(","):
        text = js_trim(part)
        type_match = _AS_TYPE_RE.search(text)
        name = _WHITESPACE_CHAR_RE.split(_PASSING_PREFIX_RE.sub("", text, count=1))[0]
        out.append(
            _EventParam(
                by_val=_BYVAL_RE.search(text) is not None,
                type=type_match.group(1) if type_match is not None else "Variant",
                is_array=name.endswith("()"),
            )
        )
    return out


def _mismatch(proc: ProcedureNode, event_params: str, is_enum: Callable[[str], bool]) -> str | None:
    """Why the handler differs from the event, or None when it matches."""
    if proc.proc_kind is not ProcKind.SUB:
        return "an event handler is a Sub"
    expected = _parse_event_params(event_params)
    actual = proc.params
    if len(actual) != len(expected):
        plural = "" if len(expected) == 1 else "s"
        return f"the event passes {len(expected)} parameter{plural}, this Sub takes {len(actual)}"
    for i, param in enumerate(actual):
        want = expected[i]
        label = f"parameter {i + 1}, '{param.name}'"
        if param.optional or param.param_array:
            what = "Optional" if param.optional else "a ParamArray"
            return f"{label} is {what}, and the event's is not"
        if bool(param.is_array) != want.is_array:
            return (
                f"{label} is an array, and the event's is not"
                if param.is_array
                else f"{label} is not an array, and the event's is"
            )
        if param.by_val != want.by_val:
            return f"{label} must be ByVal" if want.by_val else f"{label} must be ByRef, not ByVal"
        if not _same_type(param, want.type, is_enum):
            return f"{label} is {_declared_type(param)}, and the event's is {want.type}"
    return None


def _declared_type(param: ParameterNode) -> str:
    if param.as_type is not None:
        return param.as_type
    return f"{param.type_suffix} (a type suffix)" if param.type_suffix else "Variant"


_SUFFIX_TYPES: Mapping[str, str] = {
    "%": "integer",
    "&": "long",
    "!": "single",
    "#": "double",
    "@": "currency",
    "$": "string",
}


def _same_type(param: ParameterNode, expected: str, is_enum: Callable[[str], bool]) -> bool:
    """Whether a parameter's type is the event's. A library prefix names the same
    type, and an enum - an Office or MSForms one (XlXmlExportResult, fmAction)
    or the project's own - may be declared As Long, as measured."""
    if param.as_type:
        actual = _bare_type(param.as_type)
    elif param.type_suffix:
        actual = _SUFFIX_TYPES.get(param.type_suffix, "")
    else:
        actual = "variant"
    want = _bare_type(expected)
    if actual == want:
        return True
    return actual == "long" and is_enum(expected)


def _bare_type(type_: str) -> str:
    trimmed = js_trim(type_)
    bare = trimmed[trimmed.rfind(".") + 1 :]
    normalized = normalize_type(bare)
    return normalized if normalized is not None else bare.lower()
