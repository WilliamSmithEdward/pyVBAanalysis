"""Readers: VBE header handling, the loose-file loader, and the workbook reader."""

from __future__ import annotations

import types
from pathlib import Path

import pytest

from pyvbaanalysis import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.reader import (
    LooseFileReadError,
    analyze_loose_file,
    analyze_loose_files,
    analyze_workbook,
    classify_module_kind,
    load_loose_module,
    module_name_from_text,
    read_workbook_modules,
    strip_export_header,
)
from pyvbaanalysis.reader import workbook as workbook_mod
from pyvbaanalysis.symbols import ModuleSymbolKind

# Realistic export bodies (VBE uses CRLF; the loaders normalize via splitlines).
_BAS = 'Attribute VB_Name = "Mod1"\r\nOption Explicit\r\n\r\nPublic Sub Foo()\r\nEnd Sub\r\n'
_CLS = (
    "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
    'Attribute VB_Name = "Widget"\r\nAttribute VB_Exposed = False\r\n'
    "Option Explicit\r\n\r\nPublic Sub Bar()\r\nEnd Sub\r\n"
)
_DOC = (
    "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
    'Attribute VB_Name = "ThisWorkbook"\r\n'
    'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    "Option Explicit\r\n\r\nPublic Sub W()\r\nEnd Sub\r\n"
)
_FRM = (
    "VERSION 5.00\r\n"
    "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 \r\n"
    '   Caption = "x"\r\n'
    "   Begin Forms.CommandButton b\r\n"
    '      Caption = "OK"\r\n'
    "   End\r\n"
    "End\r\n"
    'Attribute VB_Name = "UserForm1"\r\n\r\nPrivate Sub b_Click()\r\nEnd Sub\r\n'
)


# -- VBE header handling ---------------------------------------------------


def test_strip_keeps_plain_bas() -> None:
    assert strip_export_header(_BAS) == _BAS


def test_strip_removes_class_designer_block() -> None:
    body = strip_export_header(_CLS)
    assert body.lstrip().startswith('Attribute VB_Name = "Widget"')
    assert "VERSION" not in body and "BEGIN" not in body and "MultiUse" not in body


def test_strip_removes_nested_form_designer_block() -> None:
    body = strip_export_header(_FRM)
    assert body.lstrip().startswith('Attribute VB_Name = "UserForm1"')
    assert "Begin" not in body and "Caption" not in body


def test_strip_unbalanced_designer_block_keeps_body() -> None:
    # A truncated header whose Begin block never closes must not swallow the body.
    truncated = (
        "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 \r\n"
        '   Caption = "x"\r\nPrivate Sub b_Click()\r\nEnd Sub\r\n'
    )
    assert strip_export_header(truncated) == truncated


def test_classify_by_extension() -> None:
    assert classify_module_kind(_BAS, extension=".bas") is ModuleSymbolKind.STANDARD
    assert classify_module_kind(_CLS, extension=".cls") is ModuleSymbolKind.CLASS
    assert classify_module_kind(_FRM, extension=".frm") is ModuleSymbolKind.USERFORM


def test_classify_document_via_vb_base() -> None:
    assert classify_module_kind(_DOC, extension=".cls") is ModuleSymbolKind.DOCUMENT


def test_classify_workbook_class_with_vb_base_is_class() -> None:
    # Workbook reads carry a VB_Base line on class modules too, with the generic VBA
    # class base GUID; that must classify as CLASS, not DOCUMENT (only a host document
    # coclass GUID means a document module).
    class_text = (
        'Attribute VB_Name = "stdThing"\r\n'
        'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\n'
        "Attribute VB_PredeclaredId = True\r\n"
        "Option Explicit\r\nPublic Sub Go()\r\nEnd Sub\r\n"
    )
    assert classify_module_kind(class_text, pyopenvba_standard=False) is ModuleSymbolKind.CLASS
    # The same module with a real host document coclass GUID is a document.
    doc_text = class_text.replace(
        "FCFB3D2A-A0FA-1068-A738-08002B3371B5", "00020820-0000-0000-C000-000000000046"
    )
    assert classify_module_kind(doc_text, pyopenvba_standard=False) is ModuleSymbolKind.DOCUMENT


def test_classify_userform_via_designer_block_without_extension() -> None:
    assert classify_module_kind(_FRM) is ModuleSymbolKind.USERFORM


def test_classify_pyopenvba_flags() -> None:
    assert classify_module_kind(_BAS, pyopenvba_standard=True) is ModuleSymbolKind.STANDARD
    assert classify_module_kind(_CLS, pyopenvba_standard=False) is ModuleSymbolKind.CLASS


def test_module_name_from_text() -> None:
    assert module_name_from_text(_CLS, "fallback") == "Widget"
    assert module_name_from_text("Public Sub S()\nEnd Sub\n", "fallback") == "fallback"


# -- loose-file loader -----------------------------------------------------


def _write(tmp_path: Path, name: str, content: str) -> Path:
    path = tmp_path / name
    path.write_bytes(content.encode("cp1252"))
    return path


def test_load_loose_module_kinds(tmp_path: Path) -> None:
    assert load_loose_module(_write(tmp_path, "Mod1.bas", _BAS)).kind is ModuleSymbolKind.STANDARD
    assert load_loose_module(_write(tmp_path, "Widget.cls", _CLS)).kind is ModuleSymbolKind.CLASS
    assert load_loose_module(_write(tmp_path, "ThisWorkbook.cls", _DOC)).kind is ModuleSymbolKind.DOCUMENT
    assert load_loose_module(_write(tmp_path, "UserForm1.frm", _FRM)).kind is ModuleSymbolKind.USERFORM


def test_analyze_loose_file_has_no_header_diagnostics(tmp_path: Path) -> None:
    # The class header must not produce spurious statement-outside-procedure noise.
    codes = {d.code for d in analyze_loose_file(_write(tmp_path, "Widget.cls", _CLS))}
    assert "statement-outside-procedure" not in codes
    assert "option-after-declaration" not in codes


def test_analyze_loose_files_cross_module(tmp_path: Path) -> None:
    person = 'Attribute VB_Name = "Person"\r\nPublic Sub Save()\r\nEnd Sub\r\n'
    entry = (
        'Attribute VB_Name = "Module1"\r\n'
        "Public Sub S()\r\n    Dim p As Person\r\n    Set p = New Person\r\n    p.Delete\r\nEnd Sub\r\n"
    )
    paths = [_write(tmp_path, "Person.cls", person), _write(tmp_path, "Module1.bas", entry)]
    result = analyze_loose_files(paths)
    assert "member-not-found" in {d.code for d in result["Module1"]}


def test_loose_file_encoding_fallback(tmp_path: Path) -> None:
    # A CP1252 smart quote (U+2019 -> byte 0x92, invalid UTF-8) must not abort the load.
    smart_quote = chr(0x2019)
    content = f'Attribute VB_Name = "Mod1"\r\n\' don{smart_quote}t panic\r\nPublic Sub S()\r\nEnd Sub\r\n'
    path = tmp_path / "Mod1.bas"
    path.write_bytes(content.encode("cp1252"))
    module = load_loose_module(path)
    assert module.name == "Mod1" and module.kind is ModuleSymbolKind.STANDARD


def test_analyze_loose_file_suppresses_cross_module_rules(tmp_path: Path) -> None:
    # A single file in isolation must not report undeclared-variable / unknown-call for
    # symbols that may be declared in a module it cannot see.
    src = (
        'Attribute VB_Name = "Mod1"\r\nOption Explicit\r\n'
        "Public Sub S()\r\n    Call HelperElsewhere\r\n    n = GlobalElsewhere\r\nEnd Sub\r\n"
    )
    path = _write(tmp_path, "Mod1.bas", src)
    codes = {d.code for d in analyze_loose_file(path)}
    assert "unknown-call" not in codes and "undeclared-variable" not in codes
    # Opting into whole-project treatment (this file genuinely is the project) surfaces them.
    whole = {d.code for d in analyze_loose_file(path, whole_project=True)}
    assert "unknown-call" in whole


def test_whole_project_flag_gates_cross_module_rules() -> None:
    src = "Option Explicit\nPublic Sub S()\n    Call HelperElsewhere\nEnd Sub\n"
    base = {
        "module_name": "Mod1",
        "module_kind": ModuleSymbolKind.STANDARD,
        "known_procedures": frozenset({"s"}),
    }
    fires = {d.code for d in analyze_module(src, AnalyzeModuleOptions(**base, whole_project=True))}
    assert "unknown-call" in fires
    silent = {d.code for d in analyze_module(src, AnalyzeModuleOptions(**base, whole_project=False))}
    assert "unknown-call" not in silent


def test_load_loose_module_on_directory_raises_contained_error(tmp_path: Path) -> None:
    # A path that is a directory (or otherwise unreadable) must surface as a contained
    # LooseFileReadError, not a raw IsADirectoryError/PermissionError.
    (tmp_path / "amodule.bas").mkdir()
    with pytest.raises(LooseFileReadError):
        load_loose_module(tmp_path / "amodule.bas")
    with pytest.raises(LooseFileReadError):
        analyze_loose_file(tmp_path / "amodule.bas")


def test_loose_severity_override_is_case_insensitive(tmp_path: Path) -> None:
    # A mis-cased override code must still apply (codes are matched case-insensitively).
    no_explicit = 'Attribute VB_Name = "Mod1"\r\nPublic Sub S()\r\nEnd Sub\r\n'
    path = _write(tmp_path, "Mod1.bas", no_explicit)
    assert "option-explicit-missing" in {d.code for d in analyze_loose_files([path])["Mod1"]}
    overridden = analyze_loose_files([path], severity_overrides={"Option-Explicit-Missing": "off"})
    assert "option-explicit-missing" not in {d.code for d in overridden["Mod1"]}


# -- workbook reader (fake pyOpenVBA) --------------------------------------


def _fake_pyopenvba(modules: list) -> types.SimpleNamespace:  # type: ignore[type-arg]
    class _Kind:
        standard = "standard"
        other = "other"

    class _Office:
        def __init__(self, *_a: object, **_k: object) -> None:
            self._modules = modules

        def __enter__(self) -> "_Office":
            return self

        def __exit__(self, *_a: object) -> bool:
            return False

        def vba_project(self) -> types.SimpleNamespace:
            return types.SimpleNamespace(modules=self._modules)

    return types.SimpleNamespace(
        ExcelFile=_Office,
        VBAModuleKind=_Kind,
        PyOpenVBAError=Exception,
    )


def _fake_module(name: str, source: str, kind: str) -> types.SimpleNamespace:
    return types.SimpleNamespace(name=name, source=source, kind=kind)


def test_workbook_reader_maps_kinds_and_strips_headers(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = _fake_pyopenvba(
        [
            _fake_module("LibJSON", _BAS, "standard"),
            _fake_module("Widget", _CLS, "other"),
            _fake_module("ThisWorkbook", _DOC, "other"),
        ]
    )
    monkeypatch.setattr(workbook_mod, "_require_pyopenvba", lambda: fake)
    modules = {m.name: m for m in read_workbook_modules("book.xlsm")}
    assert modules["LibJSON"].kind is ModuleSymbolKind.STANDARD
    assert modules["Widget"].kind is ModuleSymbolKind.CLASS
    assert modules["ThisWorkbook"].kind is ModuleSymbolKind.DOCUMENT
    assert "VERSION" not in modules["Widget"].source  # header stripped


def test_analyze_workbook_wiring(monkeypatch: pytest.MonkeyPatch) -> None:
    person = 'Attribute VB_Name = "Person"\r\nPublic Sub Save()\r\nEnd Sub\r\n'
    entry = (
        'Attribute VB_Name = "Module1"\r\n'
        "Public Sub S()\r\n    Dim p As Person\r\n    Set p = New Person\r\n    p.Delete\r\nEnd Sub\r\n"
    )
    fake = _fake_pyopenvba(
        [_fake_module("Person", person, "other"), _fake_module("Module1", entry, "standard")]
    )
    monkeypatch.setattr(workbook_mod, "_require_pyopenvba", lambda: fake)
    result = analyze_workbook("book.xlsm")
    assert "member-not-found" in {d.code for d in result["Module1"]}
    assert set(analyze_workbook("book.xlsm", only=["Module1"])) == {"Module1"}


def test_unsupported_extension_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workbook_mod, "_require_pyopenvba", lambda: _fake_pyopenvba([]))
    with pytest.raises(workbook_mod.WorkbookReadError, match="Unsupported"):
        read_workbook_modules("notes.txt")


def test_missing_pyopenvba_raises_workbook_read_error(monkeypatch: pytest.MonkeyPatch) -> None:
    import sys

    # A None entry in sys.modules makes `import pyopenvba` raise ImportError.
    monkeypatch.setitem(sys.modules, "pyopenvba", None)
    with pytest.raises(workbook_mod.WorkbookReadError, match="pyOpenVBA is required"):
        workbook_mod._require_pyopenvba()


def test_corrupt_workbook_raises_workbook_read_error(tmp_path: Path) -> None:
    # A file with an Excel extension but invalid content must be wrapped, not crash.
    bad = tmp_path / "broken.xlsm"
    bad.write_bytes(b"not a real workbook")
    with pytest.raises(workbook_mod.WorkbookReadError, match="Could not read VBA"):
        read_workbook_modules(bad)


# -- workbook reader (real Office file, optional) --------------------------

_REAL_WORKBOOK = Path(__file__).resolve().parents[2] / "xlide_vscode_testing" / "fastjson.xlsm"


@pytest.mark.skipif(not _REAL_WORKBOOK.exists(), reason="real .xlsm fixture not present")
def test_real_workbook_reads_and_classifies() -> None:
    modules = {m.name: m for m in read_workbook_modules(_REAL_WORKBOOK)}
    assert modules  # at least one module
    # Document modules (ThisWorkbook / sheets) classify as document, code modules as standard.
    assert any(m.kind is ModuleSymbolKind.DOCUMENT for m in modules.values())
    assert any(m.kind is ModuleSymbolKind.STANDARD for m in modules.values())
    # No header remnant leaks into a stripped body.
    for module in modules.values():
        assert not module.source.lstrip().upper().startswith("VERSION")
