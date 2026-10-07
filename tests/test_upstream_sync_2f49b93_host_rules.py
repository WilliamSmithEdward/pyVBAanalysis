"""Host-argument rules synced from XLIDE 67b844d..2f49b93 (hostArguments.ts and the
new hostPropertyValues.ts, worksheetFunctionArguments.ts, excelSessionState.ts,
documentNames.ts, formContents.ts and accessData.ts).

The original messages and spans came from upstream at 2f49b93. Mutable Office
state expectations follow the conservative checks in XLIDE 11.1.0.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.symbols.sheet_changes import SheetChanges, WorkbookSheetInfo
from pyvbaanalysis.symbols.symbol_model import ImplicitMember


def _sub(*lines: str) -> str:
    body = "".join(f"    {line}\n" for line in lines)
    return f"Option Explicit\nSub Main()\n{body}End Sub\n"


def _found(source: str, code: str, opts: AnalyzeModuleOptions) -> list[tuple[str, str]]:
    return [(source[d.span.start : d.span.end], d.message) for d in analyze_module(source, opts) if d.code == code]


EXCEL = AnalyzeModuleOptions(host="excel")
APP_1004 = "This will raise Run-time error '1004': Application-defined or object-defined error."


def test_a_range_chain_followed_off_the_sheet() -> None:
    source = _sub(
        'Range("B2").EntireRow.Offset(-2, 0).Select',
        'Range("C:D").Offset(0, 16383).Select',
        "Cells.Offset(1, 0).Select",
        'Range("B2").Cells(-1).Select',
    )
    assert _found(source, "host-argument-out-of-range", EXCEL) == [
        ("-2, 0", f"Offset(-2, 0) on Range(\"B2\").EntireRow reaches row 0, columns 1 to 16384, off the sheet. {APP_1004}"),
        (
            "0, 16383",
            f"Offset(0, 16383) on Range(\"C:D\") reaches rows 1 to 1048576, columns 16386 to 16387, off the sheet. {APP_1004}",
        ),
        ("1, 0", f"Offset(1, 0) on Cells reaches rows 2 to 1048577, columns 1 to 16384, off the sheet. {APP_1004}"),
        (
            "-1",
            f"Cells(-1) counts from the top-left cell of Range(\"B2\") and lands at row 0, above row 1. {APP_1004}",
        ),
    ]


def test_a_range_chain_thousands_of_links_long_is_followed() -> None:
    source = _sub('Range("A1")' + ".Offset(1, 0)" * 3000 + ".Offset(-3002, 0).Select")
    assert [text for text, _ in _found(source, "host-argument-out-of-range", EXCEL)] == ["-3002, 0"]


def test_addresses_names_and_multi_cell_values() -> None:
    source = _sub(
        'Range("A0").Select',
        'Range("$XFE:$XFE").Select',
        'Range(" ").Select',
        'Worksheets(1).Name = "History"',
        "Dim v As Long",
        'v = Range("A1", "B2")',
        'If Range("A1:A2") Then v = 1',
    )
    assert _found(source, "host-argument-out-of-range", EXCEL) == [
        (
            '"$XFE:$XFE"',
            "\"$XFE:$XFE\" is not a cell address Excel accepts: rows run 1 to 1048576 and columns A to XFD. "
            "This will raise Run-time error '1004': Method 'Range' of object failed.",
        ),
        (
            '" "',
            "Range takes an address or a name, and \" \" is blank. This will raise Run-time error '1004': "
            "Method 'Range' of object failed.",
        ),
    ]
    assert _found(source, "sheet-name-invalid", EXCEL) == [
        (
            '"History"',
            "Excel refuses this name: Excel keeps History for itself, in any case. This will raise Run-time "
            "error '1004': History is a reserved name.",
        )
    ]
    assert [message for _, message in _found(source, "multi-cell-range-as-scalar", EXCEL)] == [
        "Range(\"A1\", \"B2\") read as a value is a two-dimensional array, which a Long variable cannot hold. "
        "This will raise Run-time error '13': Type mismatch.",
        "Range(\"A1:A2\") read as a value is a two-dimensional array, which If cannot read as True or False. "
        "This will raise Run-time error '13': Type mismatch.",
    ]


def test_sheet_state_the_procedure_sets_up() -> None:
    # A statement that may activate a sheet or write to one, `Debug.Print` among
    # them, ends what is known, so each fact is read right after it is set up.
    source = _sub(
        "Dim w1 As Worksheet, w2 As Worksheet, r As Range",
        "Set w1 = ActiveSheet",
        "Set w2 = Worksheets.Add",
        'Set r = Intersect(w1.Range("A1"), w2.Range("A1"))',
        "w1.Range(Cells(1, 1), Cells(2, 2)).Value = 1",
        "w2.ShowAllData",
        'w2.Protect "pw"',
        'w2.Range("A1").Value = 1',
    )
    assert sorted(message for _, message in _found(source, "host-argument-out-of-range", EXCEL)) == [
        "Intersect takes ranges of one sheet, and 'w1' and 'w2' are different sheets. This will raise Run-time "
        "error '1004': Method 'Intersect' of object '_Global' failed.",
    ]


def test_worksheet_function_and_property_values() -> None:
    source = _sub(
        "Debug.Print WorksheetFunction.Power(-1, 0.5)",
        'Debug.Print WorksheetFunction.Match("zzz", Array("a", "b"), 0)',
        'Debug.Print WorksheetFunction.Match("B", Array("a", "b"), 0)',
        'Range("A1").Font.Size = 500',
    )
    assert [message for _, message in _found(source, "host-argument-out-of-range", EXCEL)] == [
        "WorksheetFunction.Power: Power(-1, 0.5) raises a negative number to a fractional power. The worksheet "
        "error is raised as Run-time error '1004': Unable to get the Power property of the WorksheetFunction class.",
        "WorksheetFunction.Match: an exact Match finds \"zzz\" nowhere in the array. The worksheet error is raised "
        "as Run-time error '1004': Unable to get the Match property of the WorksheetFunction class.",
    ]
    assert _found(source, "host-property-value-out-of-range", EXCEL) == [
        (
            "500",
            "Font.Size takes 1 to 409.5; 500 is outside that. This will raise Run-time error '1004': Unable to "
            "set the Size property of the Font class.",
        )
    ]


def test_excel_session_state() -> None:
    source = _sub(
        "Application.CutCopyMode = False",
        'Range("A1").PasteSpecial',
        "Dim w1 As Worksheet, w2 As Worksheet",
        "Set w1 = ThisWorkbook.Sheets.Add",
        "Set w2 = ThisWorkbook.Sheets.Add",
        'w1.Name = "Aa"',
        'w2.Name = "aa"',
    )
    assert _found(source, "paste-with-nothing-copied", EXCEL) == []
    assert _found(source, "sheet-name-invalid", EXCEL) == [
        (
            '"aa"',
            "\"aa\" is the name the code gave 'w1', another sheet it added to ThisWorkbook, and sheet names ignore case. This "
            "will raise Run-time error '1004': That name is already taken.",
        )
    ]


def test_a_new_word_document_and_its_names() -> None:
    source = _sub(
        "Dim d As Document, x As Variant",
        "Set d = Documents.Add",
        'd.Content.Text = "One." & vbCr & "Two."',
        "x = d.Paragraphs(3).Range.Text",
        'd.Variables.Add "zq", 1',
        'd.Variables.Add "ZQ", 1',
    )
    assert [message for _, message in _found(source, "host-argument-out-of-range", AnalyzeModuleOptions(host="word"))] == [
        "The variable \"ZQ\" was already added to d, and the names ignore case. This will raise Run-time error "
        "'5903': The Variable name already exists.",
    ]


def test_slides_a_new_presentation_holds() -> None:
    source = _sub(
        "Dim p As Presentation, a As Slide, b As Slide",
        "Set p = Presentations.Add",
        "Set a = p.Slides.Add(1, ppLayoutBlank)",
        "Set b = p.Slides.Add(2, ppLayoutBlank)",
        'a.Name = "Slide2"',
        "p.Slides.Add 5, ppLayoutBlank",
    )
    assert _found(source, "host-argument-out-of-range", AnalyzeModuleOptions(host="powerpoint")) == []


def test_access_sql_and_recordsets() -> None:
    source = _sub(
        "Dim db As DAO.Database, rs As DAO.Recordset",
        "Set db = CurrentDb",
        'db.Execute "INSERT INTO T1 SET A = 1"',
        'Debug.Print DCount("A", "T1", "A = 1 AND")',
        'Set rs = db.OpenRecordset("T1")',
        'rs!Nm = "a"',
    )
    access = AnalyzeModuleOptions(host="access")
    assert [message for _, message in _found(source, "runtime-argument-value", access)] == [
        "The INSERT has neither VALUES nor a SELECT. This will raise Run-time error '3134': Syntax error in "
        "INSERT INTO statement.",
        "The criteria of DCount end in 'AND' with nothing after it. This will raise Run-time error '3075': "
        "Syntax error (missing operator) in query expression.",
    ]
    assert [message for _, message in _found(source, "host-argument-out-of-range", access)] == [
        "'rs' is not being edited: no Edit or AddNew came since it was opened or last updated. This will raise "
        "Run-time error '3020': Update or CancelUpdate without AddNew or Edit."
    ]


def test_form_controls_the_designer_lists() -> None:
    source = _sub('L1.AddItem "a"', "L1.ListIndex = 5", "Dim x As Variant", "x = Mp.Pages(9).Caption")
    opts = AnalyzeModuleOptions(
        host="excel",
        implicit_members=[
            ImplicitMember("L1", "MSForms.ListBox", list_starts_empty=True),
            ImplicitMember("Mp", "MSForms.MultiPage", pages=("Page1", "Page2")),
        ],
        project_name_mentions={"l1": 1, "mp": 1},
    )
    assert _found(source, "host-property-value-out-of-range", opts) == []
    assert _found(source, "runtime-argument-value", opts) == []


def test_a_sheet_the_saved_workbook_lacks() -> None:
    source = _sub("Dim x As Variant", 'x = ThisWorkbook.Worksheets("Chart1").Name', "x = ThisWorkbook.Worksheets(3).Name")
    opts = AnalyzeModuleOptions(
        host="excel",
        workbook_sheets=[
            WorkbookSheetInfo("Sheet1", "worksheet"),
            WorkbookSheetInfo("Sheet2", "worksheet"),
            WorkbookSheetInfo("Chart1", "chartsheet"),
        ],
        project_sheet_changes=SheetChanges(),
    )
    assert _found(source, "sheet-not-in-workbook", opts) == []
