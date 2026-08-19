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
    # Pins the reason the seam exists; if this ever goes silent on its own the
    # test below stops proving anything.
    assert _codes(WORD_SOURCE) == ["member-not-found"]


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
    # The same source under Excel: four findings, every one of them false. Two
    # globals (ActiveDocument), one constant (wdOrientPortrait) and one member
    # (Selection.TypeText) all miss against the wrong host's surface.
    assert sorted(d.code for d in analyze_project(modules)["Mod1"]) == [
        "member-not-found",
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
        # Where nothing states the kind, the pair identifies document code-behind.
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


def test_the_excel_corpus_is_identical_with_and_without_the_token() -> None:
    """Every oracle case analyzed twice: no host token, then host='excel'.

    This is the differential that makes "absent means Excel" a property of the
    whole corpus rather than a claim about one code path.
    """
    checked = 0
    for case in CASES.values():
        baseline = case_codes(case)
        for module in case.modules:
            assert analyze_module(
                module.source, AnalyzeModuleOptions(host="excel", module_name=module.name)
            ) == analyze_module(module.source, AnalyzeModuleOptions(module_name=module.name))
        assert case_codes(case) == baseline
        checked += 1
    assert checked > 400


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
    # .ppt is knowingly absent from the readable set (pyOpenVBA 3.4.0 lists it but
    # reads it as a plain CFB), so it must fail as an extension, not as a parse.
    path = tmp_path / "Deck.ppt"
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
    with pytest.raises(WorkbookReadError, match="Unsupported file extension"):
        read_office_modules(path)


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
