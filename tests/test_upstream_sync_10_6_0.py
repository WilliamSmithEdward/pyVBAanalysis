"""Behaviour added by the sync to XLIDE 10.6.0.

Each test mirrors one of upstream's own, and upstream's recorded calls for them
replay through the port unchanged.

* The members a type library hides are in the host models, marked hidden, so they
  resolve. Where a Global interface answers for bare names, Application's hidden
  members stay out of bare scope.
* `Worksheets` is closed exactly as far as the `Sheets` the type library returns
  from it (XLIDE issue #79).
* A declaration after a procedure closed with the wrong End keyword is at module
  level (#81).
* A display signature's parameter list ends at its own closing parenthesis, and
  quoted text in it is opaque, so a member returning an array takes its empty
  argument list.

The read-only comparison fix (#78) is in test_rules_assignments_types.py, Mid$
(#80) in test_reference_kinds.py and the parser side of #81 in test_parser.py.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, DiagnosticSeverity
from pyvbaanalysis.diagnostics.callable_signatures import parse_runtime_display_signature
from pyvbaanalysis.host import application_member_names, get_host_members, host_object_model_for_token
from pyvbaanalysis.host.type_extensibility import excel_closed_type_names, host_type_resolves_when_compiling
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

_STD = ModuleSymbolKind.STANDARD


def _probe(body: str, *extra: ModuleInput, host: str | None = None) -> tuple[str, list]:
    """A standard module named Probe analyzed as part of a whole project."""
    source = f"Option Explicit\nPublic Sub T()\n{body}End Sub\n"
    results = analyze_project([*extra, ModuleInput("Probe", _STD, source)], host=host, referenced_hosts=[])
    return source, results["Probe"]


# -- hidden members ------------------------------------------------------------


def _member(qualified: str, name: str, token: str | None = None) -> dict:
    model = host_object_model_for_token(token) if token else None
    return next(dict(m) for m in get_host_members(qualified, model) if m["name"] == name)


def test_the_members_a_type_library_hides_are_carried_and_marked_hidden() -> None:
    # Found in ReDim's HostProbe, which compiles: ThisWorkbook.Title was "member not
    # found". The models came from the reference documentation, which leaves out
    # what the library marks hidden.
    for type_name, name in (
        ("Workbook", "Title"),
        ("Workbook", "Author"),
        ("Worksheet", "OnEntry"),
        ("Worksheet", "DisplayAutomaticPageBreaks"),
    ):
        assert _member(f"Excel.{type_name}", name).get("hidden") is True, name
    assert _member("Word.Document", "AutoSummarize", "word").get("hidden") is True
    assert _member("PowerPoint.Presentation", "HasRevisionInfo", "powerpoint").get("hidden") is True


def test_a_hidden_member_resolves_on_a_closed_surface() -> None:
    this_workbook = ModuleInput("ThisWorkbook", ModuleSymbolKind.DOCUMENT, "")
    _, found = _probe("    Debug.Print ThisWorkbook.Title\n", this_workbook)
    assert [d.code for d in found if d.code == "member-not-found"] == []
    _, found = _probe('    Dim ws As Worksheet\n    ws.OnEntry = "Handler"\n')
    assert [d.code for d in found if d.code == "member-not-found"] == []


def test_a_name_the_library_does_not_have_is_still_refused() -> None:
    _, found = _probe('    Dim ws As Worksheet\n    ws.OnEntryy = "Handler"\n')
    assert [d.code for d in found if d.code == "member-not-found"] == ["member-not-found"]


def test_a_hidden_application_member_is_not_callable_bare() -> None:
    # `Save` is a hidden method of _Application and no member of _Global, so
    # `Application.Save` compiles and a bare `Save` does not.
    assert _member("Excel.Application", "Save").get("hidden") is True
    source, found = _probe("    Application.Save\n    Save\n")
    unknown = [source[d.span.start : d.span.end] for d in found if d.code == "unknown-call"]
    assert unknown == ["Save"]


def _errors(host: str, body: str) -> list[str]:
    source, found = _probe(f"    {body}\n", host=host)
    return [
        f"{d.code}: {source[d.span.start : d.span.end]}"
        for d in found
        if d.severity is DiagnosticSeverity.ERROR
    ]


@pytest.mark.parametrize(
    ("host", "body"),
    [
        ("word", "Assistant.Visible = False"),
        ("word", "AnswerWizard.ClearFileList"),
        ("powerpoint", "Assistant.Visible = False"),
        ("powerpoint", "Debug.Print TypeName(Dialogs)"),
        ("word", "Application.ShowMe"),
    ],
)
def test_a_hidden_global_called_bare_is_accepted(host: str, body: str) -> None:
    # Word's and PowerPoint's Global interfaces hide the Office Assistant, a
    # staple of older macros, which read as "Variable not defined".
    assert _errors(host, body) == []


def test_a_hidden_application_member_global_does_not_carry_is_refused_bare() -> None:
    # ShowMe is hidden on Word's _Application and absent from _Global.
    assert _errors("word", "ShowMe") == ["unknown-call: ShowMe"]


def test_application_hidden_members_stay_in_bare_scope_only_without_a_global() -> None:
    # Access binds Application itself bare, so hidden or not, its members are
    # callable unqualified. Only where a Global interface answers do Application's
    # hidden members stay out.
    members = [{"name": "Visible", "kind": "property"}, {"name": "SecretThing", "kind": "method", "hidden": True}]

    def model(global_type: str | None = None) -> dict:
        types: dict = {"Test.Application": {"displayName": "Application", "members": members}}
        built = {
            "source": "test",
            "types": types,
            "aliases": {"application": "Test.Application"},
            "globals": {"Application": "Test.Application"},
            "constants": {},
            "memberSignatures": {},
        }
        if global_type:
            types[global_type] = {"displayName": "Global", "members": []}
            built["globalType"] = global_type
        return built

    assert application_member_names(model()) == {"visible", "secretthing"}  # type: ignore[arg-type]
    assert application_member_names(model("Test.Global")) == {"visible"}  # type: ignore[arg-type]


# -- Worksheets (#79) --------------------------------------------------------


def _member_not_found(body: str) -> list[str]:
    lines = "\r\n".join(f"    {line}" for line in body.split("\n"))
    source = f"Option Explicit\r\n\r\nPublic Sub P(ws As Worksheet, rng As Range, wb As Workbook)\r\n{lines}\r\nEnd Sub"
    opts = AnalyzeModuleOptions(module_name="M", module_kind=_STD)
    return [d.message for d in analyze_module(source, opts) if d.code == "member-not-found"]


def test_worksheets_is_closed_as_the_sheets_the_library_returns() -> None:
    # The VBE refuses this (oracle case worksheets_unknown_member_compile), whichever
    # way it is reached.
    for receiver in ("Worksheets", "Application.Worksheets", "wb.Worksheets"):
        assert _member_not_found(f"{receiver}.NoSuchMemberXyz") == [
            "Method or data member not found: 'Excel.Worksheets.NoSuchMemberXyz'."
        ], receiver


def test_worksheets_one_stays_a_worksheet_and_the_collection_keeps_its_members() -> None:
    assert _member_not_found("Worksheets(1).NoSuchMemberXyz") == [
        "Method or data member not found: 'Excel.Worksheet.NoSuchMemberXyz'."
    ]
    assert _member_not_found("Worksheets.Add\nWorksheets(1).Calculate\nDebug.Print Worksheets.Count") == []


def test_the_flag_answers_for_worksheets_as_for_sheets() -> None:
    assert host_type_resolves_when_compiling("Excel.Worksheets")
    assert host_type_resolves_when_compiling("Worksheets")
    # The list itself stays the library's own flags.
    assert "Worksheets" not in excel_closed_type_names()


# -- a procedure closed with the wrong End keyword (#81) ---------------------


def test_a_declaration_after_a_wrongly_closed_procedure_is_after_it_not_inside_it() -> None:
    # The VBE takes End Function as the closer of a Property Get, so the line after
    # it is at module level. It is still misplaced, after a procedure.
    source = "Option Explicit\nPublic Property Get P() As Long\n    P = 1\nEnd Function\nPrivate m As Long\n"
    found = analyze_module(source)
    assert [d for d in found if d.code == "module-declaration-in-procedure"] == []
    after = [d for d in found if d.code == "module-declaration-after-procedure"]
    assert [source[d.span.start : d.span.end] for d in after] == ["Private"]


# -- a member returning an array -----------------------------------------------

_THING = (
    "Public Function Values() As Long()\nEnd Function\n"
    "Public Function Pair(ByVal a As Long) As Long()\nEnd Function\n"
)


def _argument_count(thing: str, body: str) -> tuple[str, list]:
    index = ProjectIndex()
    index.set_module(ModuleInput("Thing", ModuleSymbolKind.CLASS, thing))
    source = (
        "Private Function Getter() As Thing\n    Set Getter = New Thing\nEnd Function\n"
        "Public Sub T()\n    Dim t As New Thing\n    Dim x() As Long\n    Dim n As Long\n"
        f"    {body}\nEnd Sub\n"
    )
    opts = AnalyzeModuleOptions(project_class_members=index.project_class_members())
    return source, [d for d in analyze_module(source, opts) if d.code == "argument-count"]


@pytest.mark.parametrize(
    "body",
    [
        pytest.param("x = t.Values()", id="through a variable"),
        pytest.param("n = t.Values()(0)", id="then indexed"),
        pytest.param("x = Getter().Values()", id="on a returned object"),
        pytest.param("x = Getter.Values()", id="on an implicitly called one"),
    ],
)
def test_a_member_returning_an_array_takes_its_empty_argument_list(body: str) -> None:
    # Found in vbaSQLBridge's tests, which run: `Bridge.EmptyBytes()` against
    # `Public Function EmptyBytes() As Byte()` reported "expected 1 argument".
    assert _argument_count(_THING, body)[1] == []


def test_a_missing_argument_to_a_member_returning_an_array_is_still_reported() -> None:
    source, hits = _argument_count(_THING, "x = t.Pair()")
    assert [source[d.span.start : d.span.end] for d in hits] == ["Pair"]
    assert "expected 1 argument, but got 0" in hits[0].message


def test_a_display_signature_is_read_to_its_own_closing_parenthesis() -> None:
    def params(signature: str) -> list[str]:
        return [p.name for p in parse_runtime_display_signature("F", signature).params]

    assert params("Values() As Long()") == []
    assert params("Pair(ByVal a As Long) As Long()") == ["a"]
    # Parentheses inside the list, and quoted text that holds them.
    assert params('F(ByRef a() As Byte, [s As String = ")"]) As String()') == ["a", "s"]
    assert params('F([s As String = "(,)"], [n As Long])') == ["s", "n"]
    assert params("Unclosed(a As Long") == []


def test_a_quoted_parenthesis_in_a_default_does_not_merge_two_parameters() -> None:
    thing = 'Public Function F(Optional ByVal s As String = ")", Optional n As Long) As Long()\nEnd Function\n'
    assert _argument_count(thing, "x = t.F()")[1] == []
    assert _argument_count(thing, 'x = t.F("a", 1)')[1] == []
    hits = _argument_count(thing, 'x = t.F("a", 1, 2)')[1]
    assert "expected between 0 and 2 arguments, but got 3" in hits[0].message
