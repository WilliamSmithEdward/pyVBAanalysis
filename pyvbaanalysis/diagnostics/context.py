"""Per-pass shared state for the diagnostics engine.

Ported from analysisContext.ts. RulePassContext carries the member-completion
context (member_ctx) as a first-class field: it is assembled once per pass and the
member-not-found, object-state, call-shape, type-of-is, and assignment rules read
it through ctx.member_ctx. The host object model reaches that context through
AnalyzeModuleOptions.host_model.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Literal, Protocol

from ..completion import MemberCompletionContext
from ..conditional import ConditionalActivityTracker, ConditionalCompilationEnvironment
from ..host.host_model import get_excel_object_model, get_host_members
from ..host.msforms import VBA_USERFORM_TYPE, msforms_control_members
from ..lexer.token_helpers import cached_statement_tokens
from ..lexer.token_kinds import VbaToken
from ..parser.nodes import ModuleNode, Span
from ..symbols.symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaProjectTypeName,
    VbaSymbol,
)
from .model import VbaDeclareVariableData, VbaDiagnosticData

if TYPE_CHECKING:
    from ..symbols.project_index import FormControlInfo
    from ..symbols.sheet_changes import SheetChanges, WorkbookSheetInfo
    from .opened_file_numbers import OpenedFileNumbers

__all__ = [
    "AnalysisFailure",
    "AnalyzeModuleOptions",
    "DiagnosticSeverityOverrides",
    "PushFn",
    "RulePassContext",
    "VbaDeclareVariableData",
    "is_object_module_kind",
    "own_object_member_names",
    "statement_tokens",
]

# Per-rule severity overrides keyed by stable diagnostic code; "off" disables an
# allowed rule.
DiagnosticSeverityOverrides = Mapping[str, str]


@dataclass(frozen=True, slots=True)
class AnalysisFailure:
    """Where analyze_module recovered from a failure of its own, and so what it did
    not check: `options`, an option it could not use and left out; `analysis`, the
    whole pass (nothing was checked); `rule`, one rule; `statement-walk` and
    `expression-walk`, one rule's visitor for the rest of the module, or with no
    rule named, the walk itself for the rest of the module."""

    stage: Literal["options", "analysis", "rule", "statement-walk", "expression-walk"]
    rule: str | None = None


@dataclass(slots=True)
class AnalyzeModuleOptions:
    """Inputs for analyze_module.

    document_type and host_model are typed Any to avoid importing the completion
    and host packages here; their concrete types are EventHandlerDocumentType and
    HostObjectModel. Both are read by rules: document_type drives the
    event-handler-module-scope rule, and host_model feeds the member-completion
    context and the type-name resolver.
    """

    module_name: str | None = None
    module_kind: ModuleSymbolKind | None = None
    # True when the cross-module fields below represent the COMPLETE project (every
    # module that could define a symbol visible here). False for a partial view, e.g.
    # a single file analyzed in isolation; that suppresses the rules that need the
    # whole project (undeclared-variable, unknown-call, member-not-found), since a
    # symbol declared in an unseen module is then indistinguishable from an undefined
    # one and reporting it would be a false positive.
    whole_project: bool = True
    # Honor ``'@pyvba-ignore`` suppression directives in the source. Set False to
    # report every diagnostic regardless of in-source suppression (an audit run).
    inline_suppression: bool = True
    # Return the rules' findings as upstream's analyzeModule does, without the two
    # steps XLIDE takes before it shows them (module_analysis.py): dropping a
    # runtime-error finding under On Error Resume Next, and merging findings with
    # one code and span. The differential harness sets it to compare the rules.
    raw_rule_output: bool = False
    document_type: Any = None  # EventHandlerDocumentType (from the completion package)
    # Per-rule severity overrides keyed by stable diagnostic code; "off" disables.
    severity_overrides: Mapping[str, str] | None = None
    # Lowercased procedure names callable as bare identifiers from this module.
    known_procedures: AbstractSet[str] | None = None
    # Lowercased bare identifiers visible from this module.
    known_identifiers: AbstractSet[str] | None = None
    # Exported callable signatures grouped by lowercased procedure name.
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None = None
    # Called for each failure analyze_module recovers from rather than raises
    # (XLIDE issue #178). The diagnostics it still returns are those of the rules
    # that ran; this says what was not checked, so a caller can tell a clean module
    # from one the analyzer could not fully check. A callback that raises is ignored.
    on_internal_error: Callable[[BaseException, AnalysisFailure], None] | None = None
    project_class_members: Sequence[VbaProjectClassMembers] | None = None
    # Project-defined type names (class/document/userform, user Type, Enum) visible
    # to this module, the registry the type-name resolver searches.
    project_types: Sequence[VbaProjectTypeName] | None = None
    project_visible_symbols: Sequence[VbaSymbol] | None = None
    known_non_type_names: AbstractSet[str] | None = None
    # Lowercased Private Type and Enum names of other modules, bare and as
    # `module.name`, which this module cannot use as a type (XLIDE issue #490).
    hidden_type_names: AbstractSet[str] | None = None
    # Lowercased names of every module some module in the project declares with
    # `Implements`. A module named here is an interface, so its own members are
    # declarations for an implementer to fill in rather than unfinished code.
    implemented_interfaces: AbstractSet[str] | None = None
    # Lowercased identifier-shaped words inside every string literal in the project.
    # A Private procedure named in one may be reached through Application.Run,
    # OnTime or a control's OnAction, so the unused-procedure rule treats the name
    # as used. When omitted, the module's own string literals are searched.
    project_string_literal_words: AbstractSet[str] | None = None
    # Lowercased names `Application.Run` can reach in the project (from
    # ProjectIndex.runnable_procedure_names): every Sub and Function of a standard
    # or document module, Private ones included, bare and as `module.name` (XLIDE
    # issue #243). When omitted, Application.Run is not judged.
    project_runnable_procedures: AbstractSet[str] | None = None
    # Lowercased names some module of the project may write (from
    # ProjectIndex.written_names). A Public variable none writes keeps its initial
    # value, which the runtime rules use (XLIDE issue #241). When omitted, a Public
    # variable is never taken as unchanged.
    project_written_names: AbstractSet[str] | None = None
    # How many modules of the project mention each lowercased name (from
    # ProjectIndex.name_mentions). A form's list or MultiPage that only its own
    # module names is judged from the designer's contents (XLIDE issue #315).
    project_name_mentions: Mapping[str, int] | None = None
    # The sheets of the workbook the project lives in, as saved, in tab order
    # (XLIDE issue #229). `ThisWorkbook.Sheets("name")` and `(index)` are checked
    # against them, and only when project_sheet_changes says no code in the project
    # could have made the sheet. Absent for anything but an Excel file.
    workbook_sheets: Sequence[WorkbookSheetInfo] | None = None
    # What the project's code may do to the workbook's sheets (ProjectIndex.sheet_changes).
    project_sheet_changes: SheetChanges | None = None
    # The file numbers the project's Open statements name (ProjectIndex.opened_file_numbers).
    project_opened_file_numbers: OpenedFileNumbers | None = None
    # Members the module has that no line of its own text declares: a UserForm's
    # controls, declared by the designer. Referring to one is correct VBA. Carries
    # the type as well as the name so a member lookup can resolve it. None for a
    # form means its control list is unknown, and nothing in its code-behind is
    # then called undeclared.
    implicit_members: Sequence[FormControlInfo] | None = None
    # The host class the module's DESIGNER makes it, when the caller can say: an
    # Access form's `Access.Form`. Its members belong to the module the way a
    # control does, so calling one bare or through `Me` is correct code.
    designer_class: str | None = None
    project_integer_constants: Mapping[str, str | None] | None = None
    host_model: Any = None  # HostObjectModel (from the host package)
    # Which Office host the module belongs to, as a token ("excel", "word",
    # "powerpoint", "access", ...). Resolved through the host registry when
    # host_model is not supplied directly: absent means Excel, and a named host
    # with no model means no host knowledge at all rather than Excel's.
    host: str | None = None
    # The host tokens of the libraries the project references beyond its own host,
    # in declaration order ("excel", "word", ...). A Word document referencing Excel
    # is analyzed against both. None means the reference list is UNKNOWN, which is
    # different from [] (known to reference nothing else): a loose .bas file carries
    # no reference list, so the rule that reports a missing reference stays silent
    # there rather than guess at a project it cannot see.
    referenced_hosts: Sequence[str] | None = None
    # The names of every library the project references, as its dir stream records
    # them (`Scripting`, `MSForms`). None where the project's references are not
    # known, which keeps the rules that read it silent.
    referenced_libraries: Sequence[str] | None = None
    conditional_compilation: ConditionalCompilationEnvironment | None = None
    parsed_module: ModuleNode | None = None


class PushFn(Protocol):
    """The diagnostics sink every rule reports through."""

    def __call__(
        self, rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
    ) -> None: ...


@dataclass(slots=True)
class RulePassContext:
    """Everything one diagnostics pass computes once and every rule shares.

    `member_ctx` is the member-resolution context primed with the per-pass AST and
    full-source token stream (mirrors analysisContext.ts memberCtx assembly); the
    member-not-found and object-state rules read it.
    """

    source: str
    module_name: str
    module_kind: ModuleSymbolKind
    opts: AnalyzeModuleOptions
    mod: ModuleNode
    symbols: ModuleSymbols
    activity: ConditionalActivityTracker | None
    member_ctx: MemberCompletionContext


def is_object_module_kind(module_kind: ModuleSymbolKind | None) -> bool:
    """True for the object module kinds (class, document, userform) that own a Me."""
    return module_kind in (
        ModuleSymbolKind.CLASS,
        ModuleSymbolKind.DOCUMENT,
        ModuleSymbolKind.USERFORM,
    )


# The statement-token cache lives in the lexer layer (cached_statement_tokens)
# so the call-context helpers can share it without an import cycle; this seam
# keeps the (source, Span) signature the diagnostics rules use.
def statement_tokens(source: str, span: Span) -> list[VbaToken]:
    """Significant tokens of a statement span (no comments/newlines), memoized per pass."""
    return cached_statement_tokens(source, span.start, span.end)


_NO_NAMES: frozenset[str] = frozenset()


def own_object_member_names(opts: AnalyzeModuleOptions) -> AbstractSet[str]:
    """The members of a module's own object, which its code may name bare (XLIDE
    issue #228, measured in Excel 16.0): `UsedRange` and `Shapes` in a sheet's
    module, `FullName` and `Saved` in ThisWorkbook, `Controls`, `Tag` and `Repaint`
    in a UserForm's. A sheet whose document type is unknown takes a Worksheet's and
    a Chart's, since it may be either. A module with a designer class takes that
    class's members through designer_class_member_names."""
    if opts.designer_class:
        return _NO_NAMES
    if opts.module_kind is ModuleSymbolKind.USERFORM:
        return {member["name"].lower() for member in msforms_control_members(VBA_USERFORM_TYPE) or []}
    if opts.module_kind is not ModuleSymbolKind.DOCUMENT:
        return _NO_NAMES
    model = opts.host_model if opts.host_model is not None else get_excel_object_model()
    host = (model.get("hostName") or "Excel").lower()
    name = opts.module_name.lower() if opts.module_name is not None else None
    document_type = opts.document_type
    classes: list[str]
    if document_type == "worksheet":
        classes = ["Excel.Worksheet"]
    elif document_type == "chart":
        classes = ["Excel.Chart"]
    elif document_type == "workbook":
        classes = ["Excel.Workbook"]
    elif document_type == "document":
        classes = ["Word.Document"]
    elif host == "excel":
        classes = ["Excel.Workbook"] if name == "thisworkbook" else ["Excel.Worksheet", "Excel.Chart"]
    elif host == "word" and name == "thisdocument":
        classes = ["Word.Document"]
    else:
        classes = []
    names: set[str] = set()
    for type_name in classes:
        for member in get_host_members(type_name, model):
            names.add(member["name"].lower())
    return names
