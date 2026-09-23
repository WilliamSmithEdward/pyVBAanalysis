"""Behaviour added by the sync to XLIDE 10.5.0.

Every expectation here was checked against the pinned upstream analyzer on the same
source before it was written down.

* Expression typing reaches host globals, runtime objects, VBA and host constants,
  and member expressions (`ActiveSheet.Range("A1")` is a Range).
* External constants fold into constant expressions, so `1 / vbFalse` divides by
  zero and `Left$(s, xlAbove - 1)` asks for a negative length.
* Member calls are checked for arity and argument types against the signature the
  member resolves to, on host objects and on project classes alike.
* The statements a single-line `If` carries are read by the structural rules
  (XLIDE issue #46), and the arms of one `#If` chain are alternatives while a
  repeat inside one arm is still a repeat (#58).
* A keyword before `:=` is a parameter name (`Type:=`).
* A VB6 project has its own host model, selected by the `vb6` token.
* A UserForm whose control list is authoritative proves a member absent (#26).
* The project's own types outrank the library's (#11), and a project Enum is a
  value type rather than an object receiver.
"""

from __future__ import annotations

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.host import get_vb6_object_model, host_object_model_for_token
from pyvbaanalysis.host.msforms import (
    VBA_USERFORM_TYPE,
    msforms_control_members,
    resolve_msforms_type_name,
)
from pyvbaanalysis.symbols import ImplicitMember, ModuleInput, ModuleSymbolKind

_STD = ModuleSymbolKind.STANDARD


def _messages(source: str, code: str, **kwargs: object) -> list[str]:
    opts = AnalyzeModuleOptions(**kwargs)  # type: ignore[arg-type]
    return [d.message for d in analyze_module(source, opts) if d.code == code]


def _project_messages(modules: list[ModuleInput], name: str, code: str) -> list[str]:
    return [d.message for d in analyze_project(modules)[name] if d.code == code]


# -- external constants ------------------------------------------------------

_ZERO_DIVISOR = "Expression uses '/' with a zero divisor. This will raise Run-time error '11': Division by zero."


def test_vba_and_host_constants_fold_into_a_zero_divisor() -> None:
    assert _messages("Sub T()\n    a = 1 / vbFalse\nEnd Sub\n", "division-by-zero") == [_ZERO_DIVISOR]
    assert _messages("Sub T()\n    a = 1 / (xlLandscape - 2)\nEnd Sub\n", "division-by-zero") == [
        _ZERO_DIVISOR
    ]


def test_a_qualified_host_constant_answers_only_in_its_own_host() -> None:
    source = "Sub T()\n    a = 1 / (Word.wdMainTextStory - 1)\nEnd Sub\n"
    assert _messages(source, "division-by-zero", host="word") == [_ZERO_DIVISOR]
    assert _messages(source, "division-by-zero") == []


def test_a_host_constant_folds_into_a_runtime_argument() -> None:
    assert _messages('Sub T()\n    s = Left$("abc", xlAbove - 1)\nEnd Sub\n', "runtime-argument-value") == [
        "Argument 'Length' of 'Left$' is -1; this will raise Run-time error '5': "
        "Invalid procedure call or argument."
    ]


def test_a_declared_name_shadows_the_constant() -> None:
    source = "Sub Needs(ByVal o As Object)\nEnd Sub\nSub T()\n    Dim vbFalse As Object\n    Needs vbFalse\nEnd Sub\n"
    assert _messages(source, "argument-object-type-mismatch") == []


# -- expression typing -------------------------------------------------------


def test_a_host_member_expression_has_its_return_type() -> None:
    source = 'Sub T()\n    Dim wb As Workbook\n    Set wb = ActiveSheet.Range("A1")\nEnd Sub\n'
    assert _messages(source, "assignment-object-type-mismatch") == [
        "Object assignment to 'wb' expects Workbook, but got ActiveSheet.Range(\"A1\") As "
        "Excel.Range. This object type is not compatible with Workbook."
    ]


def test_a_mixed_element_collection_yields_a_late_bound_object() -> None:
    # Sheets(i) is a Worksheet or a Chart, so any object target accepts it.
    source = 'Sub T()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("x")\nEnd Sub\n'
    assert _messages(source, "assignment-object-type-mismatch") == []


def test_a_constant_is_no_object() -> None:
    source = "Sub T()\n    Dim target As Object\n    Set target = vbFalse\n    Set target = VBA.vbFalse\nEnd Sub\n"
    assert _messages(source, "assignment-object-type-mismatch") == [
        "Object assignment to 'target' expects Object, but got vbFalse As VbTriState. "
        "An object assignment requires an object value.",
        "Object assignment to 'target' expects Object, but got VBA.vbFalse As VbTriState. "
        "An object assignment requires an object value.",
    ]


# -- single-line If and #If arms ---------------------------------------------


def test_the_statements_a_single_line_if_carries_are_read() -> None:
    assert _messages("Public Const K As Long = 1\nSub T()\n    If True Then K = 2\nEnd Sub\n", "const-assignment") == [
        "Cannot assign to constant 'K'."
    ]
    source = (
        "Sub T()\n    Dim ws As Worksheet\n    Dim r As Range\n"
        '    Set r = ActiveSheet.Range("A1")\n    If True Then Set ws = r\nEnd Sub\n'
    )
    assert _messages(source, "assignment-object-type-mismatch") == [
        "Object assignment to 'ws' expects Worksheet, but got r As Range. "
        "This object type is not compatible with Worksheet."
    ]


def test_a_repeat_inside_one_arm_is_a_repeat() -> None:
    assert _messages(
        "Public Enum E\n#If CUSTOM_FLAG Then\n    A = 1\n    A = 2\n#End If\nEnd Enum\n",
        "duplicate-enum-member",
    ) == ["Duplicate Enum member 'A' in Enum 'E'."]
    assert _messages(
        "Public Type P\n#If CUSTOM_FLAG Then\n    F As Long\n    F As String\n#End If\nEnd Type\n",
        "duplicate-type-field",
    ) == ["Duplicate field 'F' in Type 'P'."]
    assert _messages(
        "#If CUSTOM_FLAG Then\nOption Explicit\nOption Explicit\n#End If\n", "duplicate-option"
    ) == ["Duplicate Option statement; only one 'Option Explicit' is allowed per module."]
    select = (
        "Sub T()\n    Dim n As Long\n    Select Case n\n#If CUSTOM_FLAG Then\n"
        "        Case Else\n        Case Else\n#End If\n    End Select\nEnd Sub\n"
    )
    assert _messages(select, "duplicate-case-else") == ["A 'Select Case' block can have only one 'Case Else'."]


def test_the_arms_of_one_chain_are_alternatives() -> None:
    assert _messages(
        "Public Enum E\n#If CUSTOM_FLAG Then\n    A = 1\n#Else\n    A = 2\n#End If\nEnd Enum\n",
        "duplicate-enum-member",
    ) == []
    assert _messages(
        "#If CUSTOM_FLAG Then\nOption Explicit\n#Else\nOption Explicit\n#End If\n", "duplicate-option"
    ) == []


# -- member calls ------------------------------------------------------------

_PERSON = ModuleInput("Person", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Sub Save(ByVal Count As Long)\nEnd Sub\n")


def test_project_class_member_calls_are_checked() -> None:
    caller = (
        "Option Explicit\n\nPublic Sub Drive()\n    Dim p As Person\n    Set p = New Person\n"
        '    p.Save\n    p.Save "bad"\n    p.Save 1\n    With p\n        .Save "bad"\n    End With\n'
        "    Call p.Save()\nEnd Sub\n"
    )
    modules = [ModuleInput("Caller", _STD, caller), _PERSON]
    arity = "Wrong number of arguments to 'Save': expected 1 argument, but got 0."
    assert _project_messages(modules, "Caller", "argument-count") == [arity, arity]
    mismatch = (
        "Argument 'Count' of 'Save' expects Long, but got String literal \"bad\". This string "
        "literal cannot be converted to a numeric value. This will raise Run-time error '13': "
        "Type mismatch."
    )
    assert _project_messages(modules, "Caller", "argument-type-mismatch") == [mismatch, mismatch]


def test_host_member_call_argument_types_are_checked() -> None:
    assert _messages('Sub T()\n    Call Application.DeleteCustomList("bad")\nEnd Sub\n', "argument-type-mismatch") == [
        "Argument 'ListNum' of 'DeleteCustomList' expects Long, but got String literal \"bad\". "
        "This string literal cannot be converted to a numeric value. This will raise Run-time "
        "error '13': Type mismatch."
    ]


# -- VB6 ---------------------------------------------------------------------


def test_the_vb6_token_selects_the_vb6_model() -> None:
    model = host_object_model_for_token("vb6")
    assert model is get_vb6_object_model()
    assert model is not None and model["globals"]["App"] == "VB.App"


def test_a_vb6_module_knows_its_runtime_and_nothing_else() -> None:
    known = "Option Explicit\nSub T()\n    Dim s As String\n    s = App.Title\n    Debug.Print s\nEnd Sub\n"
    assert _messages(known, "undeclared-variable", host="vb6", known_identifiers=set()) == []
    unknown = "Option Explicit\nSub T()\n    Dim n As Long\n    n = vbNotAConstant\nEnd Sub\n"
    assert _messages(unknown, "undeclared-variable", host="vb6", known_identifiers=set()) == [
        "Variable not defined: 'vbNotAConstant'. Declare it before using it, or remove Option Explicit."
    ]


# -- forms (#26) -------------------------------------------------------------

_FORM_CODE = "Public Sub Accept()\nEnd Sub\n"
_FORM_HEADER_SOURCE = (
    "VERSION 5.00\n"
    "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm\n"
    '   Caption         =   "Entry"\n'
    "   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} NameBox\n"
    "   End\n"
    "End\n" + _FORM_CODE
)
_MISSING_CONTROL = 'Option Explicit\n\nPublic Sub MissingControl()\n    EntryForm.NoSuchControl.Text = "?"\nEnd Sub\n'
_NOT_FOUND = ["Method or data member not found: 'EntryForm.NoSuchControl'."]


def _form_caller(caller: str, form: ModuleInput) -> list[str]:
    return _project_messages([ModuleInput("Caller", _STD, caller), form], "Caller", "member-not-found")


def test_a_supplied_control_list_proves_a_member_absent() -> None:
    form = ModuleInput(
        "EntryForm", ModuleSymbolKind.USERFORM, _FORM_CODE,
        implicit_members=[ImplicitMember("NameBox", "MSForms.TextBox")],
    )
    assert _form_caller(_MISSING_CONTROL, form) == _NOT_FOUND


def test_a_designer_header_proves_a_member_absent() -> None:
    form = ModuleInput("EntryForm", ModuleSymbolKind.USERFORM, _FORM_HEADER_SOURCE)
    assert _form_caller(_MISSING_CONTROL, form) == _NOT_FOUND


def test_an_empty_supplied_control_list_is_still_authoritative() -> None:
    form = ModuleInput("EntryForm", ModuleSymbolKind.USERFORM, _FORM_CODE, implicit_members=[])
    assert _form_caller(_MISSING_CONTROL, form) == _NOT_FOUND


def test_a_form_whose_designer_nobody_read_proves_nothing() -> None:
    form = ModuleInput("EntryForm", ModuleSymbolKind.USERFORM, _FORM_CODE)
    assert _form_caller(_MISSING_CONTROL, form) == []


def test_the_members_a_form_has_are_never_reported() -> None:
    caller = (
        "Option Explicit\n\nPublic Sub UsesRealMembers()\n"
        '    EntryForm.NameBox.Text = "x"\n'  # designer control
        "    EntryForm.Accept\n"  # code-behind member
        "    EntryForm.Show\n"  # VBA's UserForm extender
        '    EntryForm.Caption = "t"\n'  # MSForms UserForm
        "    EntryForm.Controls.Clear\n"  # MSForms UserForm collection
        "End Sub\n"
    )
    form = ModuleInput(
        "EntryForm", ModuleSymbolKind.USERFORM, _FORM_CODE,
        implicit_members=[ImplicitMember("NameBox", "MSForms.TextBox")],
    )
    assert _form_caller(caller, form) == []


def test_the_userform_surface_carries_the_extender_members() -> None:
    members = {member["name"] for member in msforms_control_members(VBA_USERFORM_TYPE) or []}
    assert {"Show", "Hide", "Caption", "Controls"} <= members
    assert msforms_control_members("MSForms.NoSuchControlType") is None
    assert resolve_msforms_type_name("msforms . textbox") == "MSForms.TextBox"
    assert resolve_msforms_type_name("TextBox") is None


# -- project types -------------------------------------------------------------


def test_a_project_class_outranks_the_library_type_of_its_name() -> None:
    caller = "Option Explicit\n\nPublic Sub UsesFont()\n    Dim f As Font\n    Set f = New Font\n    f.Apply\n    f.Missing\nEnd Sub\n"
    font = ModuleInput("Font", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Sub Apply()\nEnd Sub\n")
    assert _project_messages([ModuleInput("Caller", _STD, caller), font], "Caller", "member-not-found") == [
        "Method or data member not found: 'Font.Missing'."
    ]


def test_a_project_enum_variable_is_no_object_receiver() -> None:
    caller = "Option Explicit\n\nPublic Sub UsesEnum()\n    Dim c As Corner\n    c = TopLeft\n    Debug.Print c.Anything\nEnd Sub\n"
    shapes = ModuleInput("Shapes2", _STD, "Public Enum Corner\n    TopLeft\n    TopRight\nEnd Enum\n")
    assert _project_messages([ModuleInput("Caller", _STD, caller), shapes], "Caller", "member-not-found") == []
