"""Behaviour changed by the sync to XLIDE 10.14.0 (issues #140 to #148).

Each issue was filed from a measurement in Excel 16.0 (build 20326) through
pyVBAharness, and every expectation here was also run through the pinned upstream
analyzer. One case per behaviour; the rules' fuller tests are upstream's, whose
recorded calls replay through the port unchanged.
"""

from __future__ import annotations

from pyvbaanalysis import analyze_project
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind


def _found(source: str, code: str) -> list[tuple[str, str]]:
    """(covered text, message) for each finding of `code`."""
    return [(source[d.span.start : d.span.end], d.message) for d in analyze_module(source) if d.code == code]


def _main(*lines: str) -> str:
    body = "".join(f"    {line}\n" for line in lines)
    return f"Option Explicit\nFunction Main() As Variant\n{body}End Function\n"


def _project_codes(*modules: tuple[str, ModuleSymbolKind, str]) -> dict[str, list[str]]:
    results = analyze_project([ModuleInput(name, kind, source) for name, kind, source in modules])
    return {name: [d.code for d in found] for name, found in results.items()}


# -- #140: one-line If tails, indexed Set targets and AddressOf ---------------------


def test_a_set_in_a_one_line_if_or_into_an_array_element_is_no_operand() -> None:
    source = (
        "Option Explicit\nPrivate mItems As Collection\n\n"
        "Public Sub TimerProc()\nEnd Sub\n\n"
        "Private Function Pointer(ByVal p As LongPtr) As LongPtr\n    Pointer = p\nEnd Function\n\n"
        "Function Main() As Long\n"
        "    Dim cols(1 To 2) As Collection, c As Collection, d As Collection, p As LongPtr\n"
        "    If mItems Is Nothing Then Set mItems = New Collection\n"
        "    Set d = New Collection\n    d.Add 5\n"
        "    If d.Count > 5 Then Set c = New Collection Else Set c = d\n"
        "    Set cols(1) = c\n"
        "    p = Pointer(AddressOf TimerProc)\n"
        "    Main = mItems.Count + cols(1).Count\n"
        "End Function\n"
    )
    assert _found(source, "collection-operand") == []
    assert _found(source, "sub-used-as-value") == []
    condition = _main("Dim c As New Collection", "If c = 1 Then Main = 2")
    assert [text for text, _ in _found(condition, "collection-operand")] == ["c"]


def test_a_range_set_in_a_one_line_if_branch_is_not_a_scalar_operand() -> None:
    branch = _main(
        "Dim ws As Worksheet, r As Range",
        "Set ws = ActiveSheet",
        'If r Is Nothing Then Set r = ws.Range("A1:P36")',
        "Main = r.Cells.Count",
    )
    assert _found(branch, "multi-cell-range-as-scalar") == []
    condition = _main("Dim ws As Worksheet", "Set ws = ActiveSheet", 'If ws.Range("A1:B2") = 1 Then Main = 2')
    assert [text for text, _ in _found(condition, "multi-cell-range-as-scalar")] == ['ws.Range("A1:B2")']


# -- #141: hex literals are signed by their width -----------------------------------


def test_hex_enum_values_and_optional_defaults_are_signed() -> None:
    source = (
        "Option Explicit\n"
        "Private Enum FileAccessFlags\n    GENERIC_READ = &H80000000\n    GENERIC_WRITE = &H40000000\n"
        "    ALL_BITS = &HFFFFFFFF\nEnd Enum\n"
        "Private Function Mask(Optional ByVal m As Long = &HFFFFFFFF, "
        "Optional ByVal i As Integer = &H8000) As Double\n    Mask = CDbl(m) + i\nEnd Function\n"
    )
    assert _found(source, "const-overflow") == []
    assert _found(source, "parameter-default-type-mismatch") == []
    wide = "Option Explicit\nPrivate Function Mask(Optional ByVal i As Integer = &H10000) As Long\n    Mask = i\nEnd Function\n"
    assert [text for text, _ in _found(wide, "parameter-default-type-mismatch")] == ["&H10000"]


# -- #142: Err.Raise takes any negative Long; On Error GoTo -1 installs nothing -----


def test_err_raise_takes_negative_numbers_and_error_does_not() -> None:
    quiet = (
        "Option Explicit\nPrivate Const E_NOINTERFACE As Long = &H80004002\nSub Main()\n"
        '    Err.Raise vbObjectError + 513, "Raised", "custom"\n'
        '    Err.Raise -1000, "Raised", "negative"\n'
        "    Err.Raise E_NOINTERFACE\nEnd Sub\n"
    )
    assert _found(quiet, "runtime-argument-value") == []
    for statement, text in (("Err.Raise 0", "0"), ("Err.Raise 65536", "65536"), ("Error -1", "-1")):
        source = f"Option Explicit\nSub Main()\n    {statement}\nEnd Sub\n"
        assert [found for found, _ in _found(source, "runtime-argument-value")] == [text]


def test_resume_next_after_on_error_goto_minus_one_is_reported() -> None:
    source = "Option Explicit\nSub Main()\n    On Error GoTo -1\n    Resume Next\nEnd Sub\n"
    assert [text for text, _ in _found(source, "resume-without-error")] == ["Resume"]


# -- #143: Print lists after Then, Else or a line number ----------------------------


def test_a_print_list_after_then_else_or_a_line_number_is_no_stray_semicolon() -> None:
    quiet = _main(
        "Dim a As Long, b As Long",
        "a = 1: b = 2",
        'Open "x.txt" For Output As #1',
        "If a = 1 Then Debug.Print a; b",
        "If a = 2 Then Debug.Print a Else Debug.Print b;",
        "If a = 1 Then Print #1, a; b",
        "If a = 1 Then Write #1, a; b",
        "Close #1",
        '10  Debug.Print "a"; "b"',
    )
    assert _found(quiet, "stray-character") == []
    stray = _main("Dim n As Long", "If n = 1 Then n = 2;")
    assert [text for text, _ in _found(stray, "stray-character")] == [";"]


# -- #144: Implements signatures and read-write properties --------------------------


def test_a_string_default_holding_a_comma_is_one_parameter() -> None:
    codes = _project_codes(
        (
            "IFoo",
            ModuleSymbolKind.CLASS,
            "Option Explicit\nPublic Function Greet(Optional ByVal sep As String = \", \") As String\nEnd Function\n",
        ),
        (
            "CFoo",
            ModuleSymbolKind.CLASS,
            "Option Explicit\nImplements IFoo\n"
            "Private Function IFoo_Greet(Optional ByVal sep As String = \", \") As String\n"
            '    IFoo_Greet = "a" & sep & "b"\nEnd Function\n',
        ),
    )
    assert "implements-member-signature" not in codes["CFoo"]


def test_a_read_write_interface_property_needs_both_accessors() -> None:
    interface = ("IShape2", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Size As Long\n")
    get_only = (
        "Option Explicit\nImplements IShape2\n"
        "Private Property Get IShape2_Size() As Long\n    IShape2_Size = 1\nEnd Property\n"
    )
    assert "implements-member-missing" in _project_codes(interface, ("CSquare2", ModuleSymbolKind.CLASS, get_only))["CSquare2"]
    both = get_only + "Private Property Let IShape2_Size(ByVal RHS As Long)\nEnd Property\n"
    assert "implements-member-missing" not in _project_codes(interface, ("CSquare3", ModuleSymbolKind.CLASS, both))["CSquare3"]


# -- #145: operator precedence and loops that leave early ---------------------------


def test_the_fold_follows_vba_operator_precedence() -> None:
    source = (
        "Option Explicit\n"
        "Private Const K1 As Long = 32000 \\ 2 * 4\n"
        "Private Const K2 As Double = 200 * 200 ^ 1\n\n"
        "Function Main() As Double\n"
        "    Dim n As Long, d As Double\n"
        "    n = 32000 \\ 2 * 4\n"
        "    d = 200 * 200 ^ 1\n"
        "    Main = K1 + K2 + n + d\n"
        "End Function\n"
    )
    assert _found(source, "const-overflow") == []
    assert _found(source, "arithmetic-overflow") == []
    mod_first = _main("Dim n As Long", "n = 1 Mod 200 * 200", "Main = n")
    assert [message.split(",")[0] for _, message in _found(mod_first, "arithmetic-overflow")] == [
        "200 (Integer) * 200 (Integer) is 40000"
    ]
    assert _found(_main("Dim d As Double", "d = -2 ^ 2", "Main = d"), "arithmetic-overflow") == []


def test_a_counter_whose_loop_can_exit_first_does_not_overflow() -> None:
    leaves = _main(
        "Dim b As Byte, i As Integer",
        "For b = 0 To 255",
        "    If b = 10 Then Exit For",
        "Next b",
        "For i = 1 To 32767",
        "    If i = 20 Then Exit For",
        "Next i",
        "Main = b + i",
    )
    assert _found(leaves, "for-counter-overflow") == []
    nested_exit_only = _main(
        "Dim b As Byte, j As Long",
        "For b = 0 To 255",
        "    For j = 1 To 3",
        "        Exit For",
        "    Next j",
        "Next b",
        "Main = b",
    )
    assert [text for text, _ in _found(nested_exit_only, "for-counter-overflow")] == ["255"]


# -- #146: Reset, calls, labels and hex file numbers --------------------------------


def test_reset_a_closing_call_and_an_error_handler_are_followed() -> None:
    source = (
        "Option Explicit\n"
        "Private Sub CloseAll()\n    Close\nEnd Sub\n\n"
        "Function Main() As Long\n"
        "    Dim p As String\n"
        '    p = Environ$("TEMP") & "\\xlide_files.txt"\n'
        "    Open p For Output As #1\n    Reset\n"
        "    Open p For Output As #1\n    CloseAll\n"
        "    Open p For Output As #1\n    Close #1\n"
        '    Open p For Output As #&H1\n    Open p & "2" For Output As #&H2\n    Close #&H1, #&H2\n'
        "    Main = Logged(p)\nEnd Function\n\n"
        "Private Function Logged(ByVal p As String) As Long\n"
        "    On Error GoTo Failed\n"
        '    Open p For Output As #1\n    Print #1, "start"\n'
        '    Err.Raise 1000, "Logged", "stop"\n    Close #1\n    Logged = 1\n    Exit Function\n'
        'Failed:\n    Print #1, "failed: " & Err.Description\n    Close #1\n    Logged = 2\nEnd Function\n'
    )
    assert _found(source, "file-already-open") == []
    assert _found(source, "file-used-after-close") == []


def test_a_print_after_reset_or_a_hex_close_is_reported() -> None:
    reset = _main('Open "x.txt" For Output As #1', "Reset", 'Print #1, "x"')
    assert [text for text, _ in _found(reset, "file-used-after-close")] == ["1"]
    hex_close = _main('Open "x.txt" For Output As #3', "Close #&H3", 'Print #3, "x"')
    assert [text for text, _ in _found(hex_close, "file-used-after-close")] == ["3"]
    hex_key = _main('Open "x.txt" For Output As #&H1', 'Open "y.txt" For Output As #1')
    assert ["#1 is still open" in message for _, message in _found(hex_key, "file-already-open")] == [True]


# -- #147: collection aliases ----------------------------------------------------


def test_collection_state_follows_set_o_equals_c() -> None:
    alias = _main(
        "Dim c As Collection, o As Collection",
        "Set c = New Collection",
        "Set o = c",
        'o.Add 1, "k"',
        "o.Add 2",
        'Main = c("k") + c(2)',
    )
    assert _found(alias, "collection-key-not-found") == []
    assert _found(alias, "collection-index-out-of-range") == []
    twice = _main(
        "Dim c As Collection, o As Collection",
        "Set c = New Collection",
        "Set o = c",
        'o.Add 1, "k"',
        'c.Add 2, "k"',
        "Main = c.Count",
    )
    assert [text for text, _ in _found(twice, "collection-key-in-use")] == ['"k"']
    untracked = _main("Dim c As Collection, v As Object", "Set c = New Collection", "Set v = c", "v.Add 1", "Main = c(1)")
    assert _found(untracked, "collection-index-out-of-range") == []


# -- #148: an identifier named constructor ------------------------------------------


def test_a_variable_named_constructor_leaves_the_module_analyzed() -> None:
    # Upstream's keyword table found Object.prototype.constructor; the port's dict
    # never did, and must keep reporting the rest of such a module.
    source = (
        "Option Explicit\nFunction Main() As Long\n    Dim constructor As Long\n"
        "    constructor = 5\n    Main = constructor\nEnd Function\n\n"
        "Sub Other()\n    Dim d As Double\n    d = 10 / 0\nEnd Sub\n"
    )
    assert [text for text, _ in _found(source, "division-by-zero")] == ["0"]


# -- 10.14.1, #152: a Property Set's value is never compared with the Get ------


def test_a_set_beside_a_get_of_another_type_compiles() -> None:
    source = (
        "Option Explicit\nPrivate mItem As Object\n"
        "Public Property Get Item() As Variant\n    Set Item = mItem\nEnd Property\n"
        "Public Property Set Item(ByVal v As Object)\n    Set mItem = v\nEnd Property\n"
    )
    assert _found(source, "property-accessor-signature-mismatch") == []
    let_mismatch = (
        "Option Explicit\nPublic Property Get Size() As Long\nEnd Property\n"
        "Public Property Let Size(ByVal v As Integer)\nEnd Property\n"
    )
    assert len(_found(let_mismatch, "property-accessor-signature-mismatch")) == 1
