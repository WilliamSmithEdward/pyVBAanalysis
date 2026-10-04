"""Reference and saved sheet metadata reaches the Office analyzer."""

import io
import struct
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from pyvbaanalysis.reader import workbook
from pyvbaanalysis.reader.workbook_sheets import _biff_sheets, _zip_sheets
from pyvbaanalysis.symbols.sheet_changes import WorkbookSheetInfo


def test_reference_names_preserve_unknown_and_unicode() -> None:
    assert workbook._reference_library_names([]) == []
    assert workbook._reference_library_names([SimpleNamespace(libid="unknown")]) is None
    assert workbook._reference_library_names([SimpleNamespace(name="Scripting", name_unicode="")]) == ["Scripting"]
    assert workbook._reference_library_names([SimpleNamespace(name="ansi", name_unicode="Unicode")]) == ["Unicode"]


def test_office_analysis_passes_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    sheets = [WorkbookSheetInfo("Main", "worksheet")]
    project = workbook.OfficeProject([], "excel", [], referenced_libraries=["Scripting"], workbook_sheets=sheets)
    monkeypatch.setattr(workbook, "read_office_project", lambda path: project)
    captured = {}

    def analyze(inputs, **options):
        captured.update(options)
        return {}

    monkeypatch.setattr(workbook, "analyze_project", analyze)
    assert workbook.analyze_office_file("book.xlsm") == {}
    assert captured["referenced_libraries"] == ["Scripting"]
    assert captured["workbook_sheets"] is sheets


def test_ooxml_sheets_keep_order_kinds_and_xml_names() -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/workbook.xml", '<workbook xmlns:r="urn:rel"><sheets><sheet name="A&amp;B" r:id="r1"/><sheet name="Chart" r:id="r2"/></sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="chartsheets/sheet1.xml"/></Relationships>')
    with zipfile.ZipFile(buffer) as archive:
        assert _zip_sheets(archive, False) == [WorkbookSheetInfo("A&B", "worksheet"), WorkbookSheetInfo("Chart", "chartsheet")]


def test_biff_sheet_names_and_kinds() -> None:
    def record(kind, body):
        return struct.pack("<HH", kind, len(body)) + body

    data = record(0x85, b"\0" * 5 + b"\x02\x05\0Chart") + record(0xA, b"")
    assert _biff_sheets(data) == [WorkbookSheetInfo("Chart", "chartsheet")]


def test_xlsb_sheet_names_and_kinds() -> None:
    def wide_string(text: str) -> bytes:
        encoded = text.encode("utf-16-le")
        return struct.pack("<I", len(encoded) // 2) + encoded

    body = b"\0" * 8 + wide_string("r1") + wide_string("Chart")
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/workbook.bin", b"\x9c\x01" + bytes([len(body)]) + body)
        archive.writestr("xl/_rels/workbook.bin.rels", '<Relationships><Relationship Id="r1" Target="chartsheets/sheet1.bin"/></Relationships>')
    with zipfile.ZipFile(buffer) as archive:
        assert _zip_sheets(archive, True) == [WorkbookSheetInfo("Chart", "chartsheet")]


@pytest.mark.parametrize("suffix", [".xlsm", ".xlsb", ".xls"])
def test_saved_sheet_fixture_metadata(suffix: str) -> None:
    fixture = Path(__file__).resolve().parents[1] / "artifacts/analyzer-pin/2f49b93/tests/fixtures/binaries" / ("SheetsFixture" + suffix)
    if not fixture.is_file():
        pytest.skip("Pinned upstream Office fixtures are not present")
    project = workbook.read_office_project(fixture)
    assert project.workbook_sheets == [
        WorkbookSheetInfo("Budget", "worksheet"),
        WorkbookSheetInfo("Drawn", "worksheet"),
        WorkbookSheetInfo("Trend", "chartsheet"),
        WorkbookSheetInfo("Later", "worksheet"),
        WorkbookSheetInfo("Hidden", "worksheet"),
    ]
    assert project.referenced_libraries is not None
    assert project.referenced_libraries == ["stdole", "Office", "MSForms"]
