"""The host-model seam: which Office host a module's VBA answers to.

Ports xlide_vscode tests/vbaHostSeam.test.ts and tests/vbaHostModels.test.ts
(XLIDE issues #24/#25). The seam's semantics are deliberately asymmetric:

* ABSENT means Excel, so every caller that predates the seam is unchanged. The
  corpus differential below pins that as a property, not an anecdote.
* A NAMED host with no model means NO host knowledge, an empty model rather than
  Excel's. Telling Word's ThisDocument it has Cells and Range is the false
  positive the seam exists to remove.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from oracle_support import CASES, case_codes  # type: ignore[attr-defined]

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.host import (
    EMPTY_HOST_MODEL,
    application_member_names,
    host_object_model_for_token,
    host_token_for_file_name,
    resolve_host_constant,
    resolve_host_global,
)
from pyvbaanalysis.host.host_model import get_host_members
from pyvbaanalysis.reader import (
    WorkbookReadError,
    analyze_office_file,
    read_office_modules,
    read_office_project,
    read_workbook_modules,
)
from pyvbaanalysis.reader.vbe_module import classify_module_kind
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

# Legal Word VBA. Under Excel's exhaustive Range model `Selection.TypeText`
# reports member-not-found; under Word it is exactly right.
WORD_SOURCE = """Option Explicit

Sub FormatDoc()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    rng.Font.Bold = True
    Selection.TypeText Text:="hello"
    If ActiveDocument.PageSetup.Orientation = wdOrientPortrait Then
        MsgBox "portrait"
    End If
End Sub
"""


def _codes(source: str, **kwargs: object) -> list[str]:
    return [d.code for d in analyze_module(source, AnalyzeModuleOptions(**kwargs))]  # type: ignore[arg-type]


# -- token resolution ------------------------------------------------------


def test_absent_host_resolves_to_the_excel_default() -> None:
    assert host_object_model_for_token(None) is None
    assert host_object_model_for_token("excel") is None
    assert host_object_model_for_token("Excel") is None
    assert host_object_model_for_token("") is None


@pytest.mark.parametrize("token", ["word", "powerpoint", "access"])
def test_a_modelled_host_answers_its_own_model(token: str) -> None:
    model = host_object_model_for_token(token)
    assert model is not None and model is not EMPTY_HOST_MODEL
    assert token in model["source"].lower()
    assert len(model["types"]) > 100


@pytest.mark.parametrize("token", ["outlook", "visio", "project", "other"])
def test_a_named_host_with_no_model_asserts_nothing(token: str) -> None:
    # Not Excel's model: an unmodelled host knows nothing rather than the wrong thing.
    assert host_object_model_for_token(token) is EMPTY_HOST_MODEL
    assert not EMPTY_HOST_MODEL["types"]
    assert not EMPTY_HOST_MODEL["globals"]


def test_an_explicit_model_outranks_the_token() -> None:
    word = host_object_model_for_token("word")
    opts = AnalyzeModuleOptions(host="excel", host_model=word)
    assert analyze_module(WORD_SOURCE, opts) == []


# -- the false positive the seam removes -----------------------------------


def test_word_source_false_positives_under_excel() -> None:
    """Pins the reason the seam exists; if this ever goes silent on its own the
    test below stops proving anything.

    A project pass is where it shows. Standalone, `Selection.TypeText` no longer
    reports under Excel at all: XLIDE 10.x treats Excel.Range as extensible, so an
    absent member on it is not provable, and the undeclared-variable findings need
    the identifier set a project pass supplies.
    """
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, WORD_SOURCE)]
    assert [d.code for d in analyze_project(modules)["Mod1"]]


@pytest.mark.parametrize("host", ["word", "outlook"])
def test_word_source_is_silent_off_the_excel_host(host: str) -> None:
    # Word's own model resolves the members; an unmodelled host asserts nothing.
    # Either way the false positive is gone.
    assert _codes(WORD_SOURCE, host=host) == []


@pytest.mark.parametrize("host", ["outlook", "visio", "project", "other"])
def test_an_unmodelled_host_reports_nothing_it_cannot_know(host: str) -> None:
    """An empty model is not uniformly quieter, and this is the case that proves it.

    Member lookups go silent (no type resolves), but the rules that ask "is this
    bare name legal" would answer no for every host-injected global, turning an
    Outlook project's own surface into a wall of undeclared-variable findings. A
    named host with no model cannot answer the question, so those rules stay
    silent for the same reason they do on a partial project view.
    """
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, WORD_SOURCE)]
    assert analyze_project(modules, host=host) == {"Mod1": []}


def test_the_host_gate_only_silences_the_rules_that_need_the_model() -> None:
    """The gate must not become a blanket mute: everything a host cannot affect
    still reports under an unmodelled host."""
    source = "Sub S()\n    Dim x As Long\n    x = 1 / 0\nEnd Sub\n"
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, source)]
    assert [d.code for d in analyze_project(modules, host="outlook")["Mod1"]] == [
        d.code for d in analyze_project(modules)["Mod1"]
    ]
    assert analyze_project(modules, host="outlook")["Mod1"]


# -- host isolation --------------------------------------------------------


@pytest.mark.parametrize(
    ("constant", "host", "value", "type_name"),
    [
        ("xlLandscape", None, 2, "XlPageOrientation"),
        ("wdMainTextStory", "word", 1, "WdStoryType"),
        ("ppLayoutBlank", "powerpoint", 12, "PpSlideLayout"),
        ("acForm", "access", 2, "AcObjectType"),
    ],
)
def test_a_host_constant_folds_on_its_own_host(
    constant: str, host: str | None, value: int, type_name: str
) -> None:
    resolved = resolve_host_constant(constant, host_object_model_for_token(host))
    assert resolved is not None
    assert resolved["value"] == value
    assert resolved["type"] == type_name


@pytest.mark.parametrize(
    ("constant", "host"),
    [
        ("xlLandscape", "word"),
        ("wdMainTextStory", None),
        ("ppLayoutBlank", "word"),
        ("acForm", None),
    ],
)
def test_a_host_constant_does_not_bleed_to_another_host(constant: str, host: str | None) -> None:
    assert resolve_host_constant(constant, host_object_model_for_token(host)) is None


@pytest.mark.parametrize(
    ("global_name", "host", "expected"),
    [
        ("ThisWorkbook", None, "Excel.Workbook"),
        ("ThisDocument", "word", "Word.Document"),
        ("Selection", "word", "Word.Selection"),
        ("ActivePresentation", "powerpoint", "PowerPoint.Presentation"),
        ("DoCmd", "access", "Access.DoCmd"),
    ],
)
def test_injected_globals_type_per_host(global_name: str, host: str | None, expected: str) -> None:
    model = host_object_model_for_token(host)
    assert resolve_host_global(global_name, model) == expected
    # Each injected global names a type the generated metadata actually carries.
    assert get_host_members(expected, model)


def test_application_members_are_injected_per_model() -> None:
    excel = application_member_names(None)
    word = application_member_names(host_object_model_for_token("word"))
    # Volatile is an Excel Application member, so a bare `Volatile` call is known
    # under Excel and unknown under Word.
    assert "volatile" in excel
    assert "volatile" not in word
    # An unmodelled host injects nothing at all.
    assert application_member_names(EMPTY_HOST_MODEL) == frozenset()


def test_me_types_host_correctly_in_a_document_module() -> None:
    from pyvbaanalysis.diagnostics.analyze_module import _me_host_type_for

    doc = ModuleSymbolKind.DOCUMENT
    assert _me_host_type_for("ThisWorkbook", doc, None) == "Excel.Workbook"
    assert _me_host_type_for("ThisWorkbook", doc, "excel") == "Excel.Workbook"
    assert _me_host_type_for("ThisDocument", doc, "word") == "Word.Document"
    # No cross-host bleed, and nothing asserted where the surface is unmodelled.
    assert _me_host_type_for("ThisWorkbook", doc, "word") is None
    assert _me_host_type_for("ThisDocument", doc, None) is None
    assert _me_host_type_for("ThisDocument", doc, "powerpoint") is None
    assert _me_host_type_for("ThisWorkbook", ModuleSymbolKind.CLASS, None) is None


# -- the container implies the host ----------------------------------------


@pytest.mark.parametrize(
    ("file_name", "expected"),
    [
        ("Book.xlsm", "excel"),
        ("Book.xlsb", "excel"),
        ("Addin.xlam", "excel"),
        ("Legacy.xls", "excel"),
        ("Report.docm", "word"),
        ("Template.dotm", "word"),
        ("Legacy.doc", "word"),
        ("Deck.pptm", "powerpoint"),
        ("Legacy.ppt", "powerpoint"),
        ("Db.accdb", "access"),
        ("Db.mdb", "access"),
        ("notes.txt", None),
        ("no-extension", None),
    ],
)
def test_a_container_implies_its_host(file_name: str, expected: str | None) -> None:
    assert host_token_for_file_name(file_name) == expected


def test_the_host_token_survives_a_project_pass() -> None:
    """A project pass is the strict case: it supplies the identifier set, so the
    undeclared-variable rule is live and Word's own constants must resolve."""
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, WORD_SOURCE)]
    assert analyze_project(modules, host="word") == {"Mod1": []}
    # The same source under Excel: three findings, every one of them false. Two
    # globals (ActiveDocument) and one constant (wdOrientPortrait) miss against the
    # wrong host's surface. Selection.TypeText used to be a fourth; Excel.Range is
    # extensible as of XLIDE 10.x, so an absent member on it is no longer reported.
    assert sorted(d.code for d in analyze_project(modules)["Mod1"]) == [
        "undeclared-variable",
        "undeclared-variable",
        "undeclared-variable",
    ]


# -- document modules beyond Excel's CLSIDs --------------------------------


def test_word_this_document_classifies_as_a_document_module() -> None:
    # Word's ThisDocument declares VB_Base = "1Normal.ThisDocument", naming no
    # CLSID, so the PredeclaredId + Exposed pair is what identifies it.
    source = (
        'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n'
        'Attribute VB_Name = "ThisDocument"\n'
        'Attribute VB_Base = "1Normal.ThisDocument"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = True\n"
        "Attribute VB_Exposed = True\n"
        "Option Explicit\n"
    )
    assert classify_module_kind(source) is ModuleSymbolKind.DOCUMENT


_PREDECLARED_EXPOSED = (
    'Attribute VB_Name = "X"\n'
    'Attribute VB_Base = "1Normal.ThisDocument"\n'
    "Attribute VB_PredeclaredId = True\n"
    "Attribute VB_Exposed = True\n"
)


@pytest.mark.parametrize(
    ("extension", "pyopenvba_standard", "expected"),
    [
        # A .bas, or the container calling it standard, says outright that the
        # module is standard; neither may be overridden by the attribute pair.
        ("bas", None, ModuleSymbolKind.STANDARD),
        (None, True, ModuleSymbolKind.STANDARD),
        # Where nothing states the kind, the pair beside a base identifies
        # document code-behind.
        ("cls", None, ModuleSymbolKind.DOCUMENT),
        (None, False, ModuleSymbolKind.DOCUMENT),
        (None, None, ModuleSymbolKind.DOCUMENT),
    ],
)
def test_a_direct_standard_signal_outranks_the_attribute_pair(
    extension: str | None, pyopenvba_standard: bool | None, expected: ModuleSymbolKind
) -> None:
    source = _PREDECLARED_EXPOSED
    assert (
        classify_module_kind(
            source, extension=extension, pyopenvba_standard=pyopenvba_standard
        )
        is expected
    )


@pytest.mark.parametrize(
    ("predeclared", "exposed"), [("True", "False"), ("False", "True"), ("False", "False")]
)
def test_the_attribute_pair_needs_both_halves(predeclared: str, exposed: str) -> None:
    source = (
        'Attribute VB_Name = "X"\n'
        f"Attribute VB_PredeclaredId = {predeclared}\n"
        f"Attribute VB_Exposed = {exposed}\n"
    )
    assert classify_module_kind(source, extension="cls") is ModuleSymbolKind.CLASS


@pytest.mark.parametrize(("extension", "pyopenvba_standard"), [("cls", None), (None, False)])
def test_the_attribute_pair_without_a_base_is_a_class(
    extension: str | None, pyopenvba_standard: bool | None
) -> None:
    # A .cls export has no VB_Base line, and a class can be predeclared and
    # exposed: '@PredeclaredId and '@Exposed written through XLIDE, or an
    # add-in's factory class.
    source = 'Attribute VB_Name = "X"\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\n'
    assert (
        classify_module_kind(source, extension=extension, pyopenvba_standard=pyopenvba_standard)
        is ModuleSymbolKind.CLASS
    )


def test_a_predeclared_exposed_class_on_the_class_base_is_a_class() -> None:
    # stdVBA's stdCOM and stdWebSocket as read out of a workbook: the VBE's
    # class base, PredeclaredId and Exposed. The base says class outright.
    source = (
        'Attribute VB_Name = "stdCOM"\n'
        'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = True\n"
        "Attribute VB_Exposed = True\n"
    )
    assert classify_module_kind(source, pyopenvba_standard=False) is ModuleSymbolKind.CLASS
    assert classify_module_kind(source) is ModuleSymbolKind.CLASS


def test_an_ordinary_class_module_is_not_a_document() -> None:
    source = (
        'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n'
        'Attribute VB_Name = "Widget"\n'
        'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\n'
        "Attribute VB_PredeclaredId = False\n"
        "Attribute VB_Exposed = False\n"
    )
    assert classify_module_kind(source) is ModuleSymbolKind.CLASS


# -- absent-host behavior is provably unchanged ----------------------------


def _is_compare_database(diagnostic: object) -> bool:
    return (
        getattr(diagnostic, "code", "") == "invalid-option-statement"
        and "Option Compare Database" in getattr(diagnostic, "message", "")
    )


def test_the_excel_corpus_is_identical_with_and_without_the_token() -> None:
    """Every oracle case analyzed twice: no host token, then host='excel'.

    This is the differential that makes "absent means Excel" a property of the
    whole corpus rather than a claim about one code path.

    One finding deliberately depends on whether a host was NAMED rather than which
    object model answers: `Option Compare Database` is an Access directive, and it
    is reported only where a project names a host that is not Access. A file no
    project claims names no host and is left alone. So that finding is set aside
    here, and the test below pins that it is the only difference.
    """
    checked = 0
    for case in CASES.values():
        baseline = case_codes(case)
        for module in case.modules:
            named = analyze_module(module.source, AnalyzeModuleOptions(host="excel", module_name=module.name))
            absent = analyze_module(module.source, AnalyzeModuleOptions(module_name=module.name))
            assert [d for d in named if not _is_compare_database(d)] == absent
        assert case_codes(case) == baseline
        checked += 1
    assert checked > 400


def test_option_compare_database_is_the_one_finding_a_named_host_adds() -> None:
    source = "Option Explicit\nOption Compare Database\n"
    assert analyze_module(source, AnalyzeModuleOptions()) == []
    named = analyze_module(source, AnalyzeModuleOptions(host="excel"))
    assert [d.code for d in named] == ["invalid-option-statement"]
    assert analyze_module(source, AnalyzeModuleOptions(host="access")) == []


# -- reading real containers -----------------------------------------------

pyopenvba = pytest.importorskip("pyopenvba")

_WORD_MODULE = """Option Explicit

Sub Greet()
    Selection.TypeText Text:="hello"
    If ActiveDocument.PageSetup.Orientation = wdOrientPortrait Then
        ActiveDocument.Save
    End If
End Sub
"""


def _word_document(tmp_path: Path) -> Path:
    """A real .docm authored by pyOpenVBA, carrying one Word module."""
    path = tmp_path / "Fixture.docm"
    with pyopenvba.WordFile.create_new(path) as document:
        document.set_module("Module1", _WORD_MODULE)
        document.save()
    return path


_FORM_CODE = (
    "Option Explicit\r\n"
    "Private Sub UserForm_Initialize()\r\n"
    '    RegionPick.AddItem "West"\r\n'
    '    Inner.Text = "x"\r\n'
    '    Me.Caption = "Pick"\r\n'
    '    Page1.Caption = "First"\r\n'
    "    Tabs.Value = 0\r\n"
    "End Sub\r\n"
    "Public Sub Dismiss()\r\n"
    "    Me.Hide\r\n"
    "End Sub\r\n"
)


def _workbook_with_form(tmp_path: Path) -> Path:
    """A real .xlsm with a UserForm: a ComboBox, a TextBox inside a Frame, and a
    MultiPage with the two pages the designer gives a new one."""
    path = tmp_path / "FormBook.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        form = book.add_form("FrmPick")
        form.add_control("ComboBox", "RegionPick")
        form.add_control("Frame", "Box")
        form.add_control("TextBox", "Inner", container="Box")
        form.add_control("MultiPage", "Tabs")
        header = book.vba_project().get_module("FrmPick").source
        book.set_module("FrmPick", header + _FORM_CODE)
        book.set_module("Module1", "Option Explicit\r\nSub Go()\r\n    FrmPick.Show\r\nEnd Sub\r\n")
        book.save()
    return path


def test_a_userform_in_a_container_is_read_as_a_form_with_its_controls(tmp_path: Path) -> None:
    by_name = {module.name: module for module in read_office_modules(_workbook_with_form(tmp_path))}
    form = by_name["FrmPick"]
    assert form.kind is ModuleSymbolKind.USERFORM
    # A MultiPage's pages are MSForms.Page to VBA (pyOpenVBA reads their site
    # class as a Form), and its internal TabStrip has no name to reach it by.
    assert [(m.name, m.type) for m in form.implicit_members or ()] == [
        ("RegionPick", "MSForms.ComboBox"),
        ("Box", "MSForms.Frame"),
        ("Inner", "MSForms.TextBox"),
        ("Tabs", "MSForms.MultiPage"),
        ("Page1", "MSForms.Page"),
        ("Page2", "MSForms.Page"),
    ]


def test_a_userform_in_a_container_analyzes_clean(tmp_path: Path) -> None:
    """Its controls, nested ones included, and the UserForm's own members
    (Caption, Hide, Show) all resolve: nothing here is undeclared or missing."""
    results = analyze_office_file(_workbook_with_form(tmp_path))
    assert results["FrmPick"] == []
    assert results["Module1"] == []


_ACCESS_FORM_CODE = (
    "Option Compare Database\r\n"
    "Option Explicit\r\n"
    "Private Sub Form_Load()\r\n"
    '    Me.Caption = "Orders"\r\n'
    "    Me.Qty.Value = 1\r\n"
    "    Qty.Locked = False\r\n"
    "    Me.Order_Date.Value = Date\r\n"
    "    Detail.Visible = True\r\n"
    "End Sub\r\n"
)


def _database_with_form(tmp_path: Path) -> Path:
    """A real .accdb with a form: a TextBox, a TextBox whose name has a space,
    and code behind it."""
    path = tmp_path / "Orders.accdb"
    db = pyopenvba.access.AccessDatabase.create_new(path)
    db.add_form("Orders")
    db.add_control("Orders", "TextBox", "Qty")
    db.add_control("Orders", "TextBox", "Order Date")
    db.set_design_code("Orders", _ACCESS_FORM_CODE)
    db.save()
    return path


def test_an_access_form_is_read_with_its_class_and_controls(tmp_path: Path) -> None:
    by_name = {module.name: module for module in read_office_modules(_database_with_form(tmp_path))}
    form = by_name["Form_Orders"]
    assert form.kind is ModuleSymbolKind.USERFORM
    assert form.designer_class == "Access.Form"
    members = {(m.name, m.type) for m in form.implicit_members or ()}
    # Access names a control for VBA by replacing what an identifier cannot
    # hold: `Order Date` is `Order_Date`.
    assert {("Qty", "Access.Textbox"), ("Order_Date", "Access.Textbox"), ("Detail", "Access.Section")} <= members


def test_an_access_form_analyzes_clean(tmp_path: Path) -> None:
    assert analyze_office_file(_database_with_form(tmp_path))["Form_Orders"] == []


# (code page, control name, module name): a form control and a module named in
# each script, read out of a project saved in that script's code page.
_NATIVE_FORM_NAMES = [
    (1251, "Имя", "МодульТест"),
    (1253, "Όνομα", "Ενότητα"),
    (1250, "Jméno", "Modul"),
    (1254, "İsimKutusu", "Modül"),
    (932, "名前", "モジュール"),
    (936, "名称", "测试模块"),
]


def _patch_project_code_page(dir_raw: bytes, code_page: int) -> bytes:
    """Rewrite the dir stream's PROJECTCODEPAGE record ([MS-OVBA] 2.3.4.2.1.4),
    as pyOpenVBA's own language matrix does. PROJECTVERSION's size slot is a
    reserved marker over a fixed 10-byte payload, so it is stepped over."""
    import struct

    buf = bytearray(dir_raw)
    pos = 0
    while pos + 6 <= len(buf):
        record_id = struct.unpack_from("<H", buf, pos)[0]
        if record_id == 0x0009:
            pos += 12
            continue
        size = struct.unpack_from("<I", buf, pos + 2)[0]
        if record_id == 0x0003:
            struct.pack_into("<H", buf, pos + 6, code_page)
            return bytes(buf)
        pos += 6 + size
    raise AssertionError("PROJECTCODEPAGE record not found")


def _workbook_in_code_page(tmp_path: Path, code_page: int, control: str, module: str) -> Path:
    import io
    import zipfile

    from pyopenvba.cfb import CFB
    from pyopenvba.vba import compress, decompress

    path = tmp_path / f"cp{code_page}.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        book.save()
    entry = "xl/vbaProject.bin"
    with zipfile.ZipFile(io.BytesIO(path.read_bytes())) as zin:
        cfb = CFB.from_bytes(zin.read(entry))
        dir_raw = decompress(cfb.get_stream_in_storage("VBA", "dir"))
        cfb.write_stream_in_storage("VBA", "dir", compress(_patch_project_code_page(dir_raw, code_page)))
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w") as zout:
            for info in zin.infolist():
                zout.writestr(info, cfb.to_bytes() if info.filename == entry else zin.read(info.filename))
    path.write_bytes(out.getvalue())
    with pyopenvba.ExcelFile(path) as book:
        form = book.add_form("Frm1")
        form.add_control("TextBox", control)
        header = book.vba_project().get_module("Frm1").source
        book.set_module(
            "Frm1",
            header + "Option Explicit\r\nPrivate Sub UserForm_Initialize()\r\n"
            f'    {control}.Text = "x"\r\n    Me.{control}.Visible = True\r\nEnd Sub\r\n',
        )
        book.vba_project().add_module(
            module,
            f'Attribute VB_Name = "{module}"\r\nOption Explicit\r\nSub Go()\r\n'
            f'    Frm1.Show\r\n    Frm1.{control}.Text = "y"\r\nEnd Sub\r\n',
            kind=pyopenvba.VBAModuleKind.standard,
        )
        book.save()
    return path


@pytest.mark.parametrize(
    ("code_page", "control", "module"), _NATIVE_FORM_NAMES, ids=[f"cp{row[0]}" for row in _NATIVE_FORM_NAMES]
)
def test_a_userform_in_a_non_western_project_reads_its_native_control_names(
    tmp_path: Path, code_page: int, control: str, module: str
) -> None:
    """A form's controls come from its designer storage, decoded in the
    project's code page: a Russian, Greek, Czech, Turkish, Japanese or Chinese
    control name reaches the analyzer intact, and code using it is clean."""
    path = _workbook_in_code_page(tmp_path, code_page, control, module)
    by_name = {m.name: m for m in read_office_modules(path)}
    assert by_name["Frm1"].kind is ModuleSymbolKind.USERFORM
    assert [(m.name, m.type) for m in by_name["Frm1"].implicit_members or ()] == [(control, "MSForms.TextBox")]
    assert module in by_name
    assert {name: found for name, found in analyze_office_file(path).items() if found} == {}


def test_an_access_class_named_in_another_case_by_the_catalog_is_a_class(tmp_path: Path) -> None:
    """The dir catalog can spell a module `Basket` while the module list reads
    `basket` (XLIDE's AccessFormFixture.accdb does). The class check must not
    depend on the case."""
    from types import SimpleNamespace

    from pyvbaanalysis.reader.workbook import _read_access_project

    class_text = (
        'Attribute VB_Name = "Basket"\r\n'
        'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\n'
        "Option Compare Database\r\nOption Explicit\r\n"
    )

    class Reader:
        def __init__(self, _path: Path) -> None:
            pass

        def __enter__(self) -> "Reader":
            return self

        def __exit__(self, *_exc: object) -> None:
            return None

        def read_project_info(self) -> SimpleNamespace:
            entry = SimpleNamespace(name="Basket", is_class_module=True)
            return SimpleNamespace(modules=[entry], references=[])

        def vba_module_names(self) -> list[str]:
            return ["basket"]

        def read_vba_module_with_attributes(self, _name: str) -> str:
            return class_text

    modules, _ = _read_access_project(SimpleNamespace(AccessReader=Reader), tmp_path / "none.accdb")
    assert [(m.name, m.kind) for m in modules] == [("basket", ModuleSymbolKind.CLASS)]


def test_a_word_container_reads_its_modules(tmp_path: Path) -> None:
    modules = read_office_modules(_word_document(tmp_path))
    by_name = {module.name: module for module in modules}
    assert "Module1" in by_name
    assert by_name["Module1"].kind is ModuleSymbolKind.STANDARD
    # Word's code-behind names no CLSID, so the PredeclaredId + Exposed pair is
    # what makes it a document module rather than a plain class.
    assert by_name["ThisDocument"].kind is ModuleSymbolKind.DOCUMENT


def test_a_word_container_analyzes_against_word(tmp_path: Path) -> None:
    """The end-to-end contract: the extension picks the host, so legal Word VBA
    read out of a .docm is silent instead of measured against Excel."""
    path = _word_document(tmp_path)
    assert analyze_office_file(path)["Module1"] == []
    # Forcing the same modules through the Excel-default project path is what the
    # seam replaced, and it is not silent.
    modules = [ModuleInput("Module1", ModuleSymbolKind.STANDARD, _WORD_MODULE)]
    assert analyze_project(modules) != {"Module1": []}


def test_the_excel_reader_points_at_the_generic_one(tmp_path: Path) -> None:
    path = _word_document(tmp_path)
    with pytest.raises(WorkbookReadError, match="analyze_office_file"):
        read_workbook_modules(path)


def test_an_unreadable_container_is_rejected_by_extension(tmp_path: Path) -> None:
    # pyOpenVBA refuses a .ppam (a PowerPoint add-in) by its extension, so it must
    # fail here as an extension, not as a confusing parse error.
    path = tmp_path / "Deck.ppam"
    path.write_bytes(b"PK\x03\x04" + b"\x00" * 64)
    with pytest.raises(WorkbookReadError, match="Unsupported file extension"):
        read_office_modules(path)


_FIXTURES = Path(__file__).parent / "fixtures"


def test_a_legacy_ppt_reads_its_modules() -> None:
    """A .ppt keeps its project in a compressed storage inside the document;
    pyOpenVBA 6 reads it. Upstream's readModules gives the same modules. The
    fixture holds VBA, so the sdist leaves it out and the test skips there."""
    fixture = _FIXTURES / "PowerPointFixture.ppt"
    if not fixture.exists():
        pytest.skip("the .ppt fixture is not in the sdist")
    project = read_office_project(fixture)
    assert project.host == "powerpoint"
    kinds = {m.name: m.kind for m in project.modules}
    assert kinds["Module1"] is ModuleSymbolKind.STANDARD
    assert kinds["CDeck"] is ModuleSymbolKind.CLASS
    assert set(kinds) >= {"ZFixtureSetup", "PyVbaHarnessRunner", "PyVbaHarnessCall"}


def test_access_module_text_keeps_its_characters_and_short_modules(tmp_path: Path) -> None:
    """Before pyOpenVBA 6.3.0, AccessReader decoded module text as latin-1 (an
    em dash came back as \\x97, a euro sign as \\x80) and could miss a short
    module altogether (pyOpenVBA #33). The floor is past that: accented names,
    and a one-line module, read intact and analyze clean."""
    path = tmp_path / "Prices.accdb"
    db = pyopenvba.access.AccessDatabase.create_new(path)
    db.set_module(
        "Module1",
        "Option Compare Database\r\nOption Explicit\r\n"
        "' Preis in € — netto\r\n"
        "Public Function Größe() As Long\r\n"
        "    Dim Zähler As Long\r\n    Zähler = 2\r\n    Größe = Zähler\r\n"
        "End Function\r\n",
    )
    db.add_module("Short", "Option Compare Database\r\n")
    db.save()
    by_name = {m.name: m for m in read_office_modules(path)}
    assert "Short" in by_name
    assert "' Preis in € — netto" in by_name["Module1"].source
    assert "Größe" in by_name["Module1"].source
    assert analyze_office_file(path)["Module1"] == []


@pytest.mark.parametrize("suffix", [".accda", ".mda"])
def test_an_access_add_in_reads_as_a_database(tmp_path: Path, suffix: str) -> None:
    # An add-in is stored as a database is, so the Access reader opens it.
    path = tmp_path / f"Tools{suffix}"
    db = pyopenvba.access.AccessDatabase.create_new(path)
    db.set_module("Module1", "Option Compare Database\r\nOption Explicit\r\nPublic Sub Go()\r\nEnd Sub\r\n")
    db.save()
    project = read_office_project(path)
    assert project.host == "access"
    assert [m.name for m in project.modules] == ["Module1"]


# -- the model memos are identity-safe -------------------------------------


def test_host_model_memos_stay_bounded_and_hold_their_keys() -> None:
    """The host memos are keyed by model identity, so they must keep the model
    alive and stay bounded. A bare dict[id(model)] does neither: the entry count
    grows once per model object ever seen, and a collected model's id can be
    recycled by a later one, which would serve one model's index for another.
    """
    import gc

    import pyvbaanalysis.host.host_model as host_model_module
    from pyvbaanalysis.host.host_model import get_host_members, is_host_member_name

    def probe(index: int) -> dict[str, object]:
        member = f"Member{index}"
        return {
            "source": f"probe {index}",
            "aliases": {},
            "globals": {},
            "types": {"P.T": {"displayName": "P.T", "members": [{"name": member, "kind": "method"}]}},
            "constants": {},
            "memberSignatures": {},
        }

    for i in range(200):
        model = probe(i)
        # Each temporary model must answer for itself, never for a predecessor
        # that happened to occupy the same address.
        assert [m["name"] for m in get_host_members("P.T", model)] == [f"Member{i}"]
        assert is_host_member_name(f"member{i}", model)
        assert not is_host_member_name("member_that_never_existed", model)
        del model
    gc.collect()

    for cache in (
        host_model_module._MODEL_INDEX_CACHE,
        host_model_module._CONSTANT_INDEX_CACHE,
        host_model_module._HOST_MEMBER_NAMES_CACHE,
        host_model_module._APPLICATION_MEMBER_NAMES,
    ):
        assert len(cache._entries) <= 8


def test_every_vendored_model_answers_independently() -> None:
    """All four models live at once, so a shared memo must not blur them."""
    models = {token: host_object_model_for_token(token) for token in ("word", "powerpoint", "access")}
    models["excel"] = None
    seen = {
        token: frozenset(m["name"].lower() for m in get_host_members(app, model))
        for token, model, app in [
            ("excel", None, "Excel.Application"),
            ("word", models["word"], "Word.Application"),
            ("powerpoint", models["powerpoint"], "PowerPoint.Application"),
            ("access", models["access"], "Access.Application"),
        ]
    }
    assert all(members for members in seen.values())
    # Interleave the lookups; a stale memo would make a later read match an earlier host.
    for token, model, app in [
        ("access", models["access"], "Access.Application"),
        ("excel", None, "Excel.Application"),
        ("word", models["word"], "Word.Application"),
    ]:
        assert frozenset(m["name"].lower() for m in get_host_members(app, model)) == seen[token]
