"""Behaviour added by the sync to XLIDE 10.12.0 (issues #97 to #133).

Every expectation here was run through the pinned upstream analyzer as well, with
the same findings, and upstream's own recorded test calls replay through the port
unchanged. One case per behaviour; the rules' fuller tests are upstream's.

* The lexer and parser read `On Local Error` (#98), `Property` as a name (#98), a
  one-line If's block opener (#128), `Else:` (#129), a block opener or closer
  restated in both arms of an `#If` (#130), and a no-break space as a stray
  character (#132).
* A bare `=` to an object variable is a Let through its default member (#107),
  an Object-declared item is late bound (#114), and a Set between two project
  interfaces one class implements compiles (#109).
* Argument, assignment, declaration, control-flow and expression rules follow the
  forms measured in Excel 16.0 (#97, #99 to #106, #108, #111 to #113, #115, #119,
  #121, #124, #125).
* New rules: overflow (#116), handler flow (#117), runtime values the code states
  (#118), array bounds (#120), collection state and Variant values (#121), host
  arguments (#122), file statements (#123), declaration forms (#124), statement
  forms and Implements members (#125), line continuations (#126), directive forms
  (#130), stray characters and line length (#132, #133), literal forms (#125,
  #133), and late-bound members.
"""

from __future__ import annotations

from pyvbaanalysis import analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.host.type_extensibility import host_type_resolves_when_compiling
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_NBSP = chr(0xA0)


def _found(source: str, code: str, opts: AnalyzeModuleOptions | None = None) -> list[tuple[str, str]]:
    """(covered text, message) for each finding of `code`."""
    return [
        (source[d.span.start : d.span.end], d.message) for d in analyze_module(source, opts) if d.code == code
    ]


def _codes(source: str, opts: AnalyzeModuleOptions | None = None) -> set[str]:
    return {d.code for d in analyze_module(source, opts)}


def _project(*modules: tuple[str, ModuleSymbolKind, str]) -> dict[str, list[tuple[str, str, str]]]:
    """Each module's findings as (code, covered text, message), analyzed as a project."""
    inputs = [ModuleInput(name, kind, source) for name, kind, source in modules]
    sources = {name: source for name, _, source in modules}
    results = analyze_project(inputs)
    return {
        name: [(d.code, sources[name][d.span.start : d.span.end], d.message) for d in found]
        for name, found in results.items()
    }


# -- the lexer and parser --------------------------------------------------------


def test_on_local_error_is_on_error_and_local_is_reserved() -> None:
    source = "Option Explicit\nSub S()\n    On Local Error Resume Next\n    On Local Error GoTo 0\nEnd Sub\n"
    assert _codes(source) == set()
    declared = "Option Explicit\nSub S()\n    Dim Local As Long\nEnd Sub\n"
    assert _found(declared, "invalid-declaration-name") == [
        ("Local", "Reserved VBA keyword 'Local' cannot be used as a variable name.")
    ]


def test_property_is_a_name_outside_a_property_header() -> None:
    source = "Option Explicit\nDim Property As Long\nSub S()\n    Property = 1\nEnd Sub\n"
    assert _codes(source) == set()


def test_else_with_a_colon_is_no_label() -> None:
    source = (
        "Option Explicit\nSub S(ByVal a As Boolean, ByVal b As Boolean)\n"
        "    If a Then\n        Debug.Print 1\n    Else:\n        Debug.Print 2\n    End If\n"
        "    If b Then\n        Debug.Print 3\n    Else:\n        Debug.Print 4\n    End If\nEnd Sub\n"
    )
    assert _codes(source) == set()


def test_a_block_opener_in_a_one_line_if_tail_opens_its_block() -> None:
    source = (
        "Option Explicit\nSub S(ByVal a As Boolean)\n    Dim c As New Collection\n"
        "    If a Then With c: .Add 1: End With\nEnd Sub\n"
    )
    assert _codes(source) == set()


def test_a_closer_restated_in_both_directive_arms_closes_one_block() -> None:
    source = (
        "Option Explicit\nSub S(ByVal a As Boolean)\n    If a Then\n        Debug.Print 1\n"
        "#If VBA7 Then\n    End If\n#Else\n    End If\n#End If\nEnd Sub\n"
    )
    assert _codes(source) == set()


def test_a_no_break_space_is_a_stray_character() -> None:
    source = f"Option Explicit\nSub S()\n    Dim n As Long\n    n = 1{_NBSP}+ 1\n    Debug.Print n\nEnd Sub\n"
    assert _found(source, "stray-character") == [
        (
            _NBSP,
            "A non-breaking space (U+00A0) is not whitespace in VBA: the VBE reads it as part of a "
            "name or refuses the line. Replace it with an ordinary space. This is a VBE compile error.",
        )
    ]


# -- object values and member access ---------------------------------------------


def test_a_bare_let_to_an_object_goes_through_its_default_member() -> None:
    source = (
        "Option Explicit\nSub S()\n    Dim r As Range, c As Collection, ws As Worksheet\n"
        "    Set r = Range(\"A1\")\n    Set c = New Collection\n    Set ws = ActiveSheet\n"
        "    r = 5\n    c = 5\n    ws = 9\nEnd Sub\n"
    )
    assert _found(source, "set-required") == [
        (
            "c",
            "Assignment to 'c' requires Set: the default member of Collection takes an argument, so a "
            "Let cannot reach it. This is a VBE compile error: Argument not optional.",
        ),
        (
            "ws",
            "Assignment to 'ws' requires Set: Worksheet has no default member for a Let to reach. This "
            "will raise Run-time error '438': Object doesn't support this property or method.",
        ),
    ]


def test_a_property_let_takes_an_object_and_set_needs_a_property_set() -> None:
    holder = (
        "Option Explicit\nPrivate m As Object\nPublic Property Let Item(ByVal v As Object)\nEnd Property\n"
        "Public Property Set Only(ByVal v As Object)\n    Set m = v\nEnd Property\n"
    )
    main = (
        "Option Explicit\nPublic Sub Main()\n    Dim h As New Holder\n    h.Item = New Collection\n"
        "    Set h.Item = New Collection\n    h.Only = New Collection\nEnd Sub\n"
    )
    found = _project(("Holder", ModuleSymbolKind.CLASS, holder), ("Module1", ModuleSymbolKind.STANDARD, main))
    assert [f for f in found["Module1"] if f[0] in ("set-required", "set-requires-object")] == [
        (
            "set-requires-object",
            "Item",
            "Set assignment to 'h.Item' needs a Property Set, but the property declares only a "
            "Property Let. This is a VBE compile error: Invalid use of property.",
        ),
        (
            "set-required",
            "Only",
            "Assignment to 'h.Only' requires Set: the property declares a Property Set and no Property "
            "Let. This is a VBE compile error: Invalid use of property.",
        ),
    ]


def test_a_set_between_interfaces_of_one_class_compiles() -> None:
    impl = (
        "Option Explicit\nImplements IFoo\nImplements IBar\nPrivate Sub IFoo_Run()\nEnd Sub\n"
        "Private Sub IBar_Halt()\nEnd Sub\n"
    )
    main = (
        "Option Explicit\nPublic Sub Main()\n    Dim a As IFoo, b As IBar, c As Impl, o As Other\n"
        "    Set a = New Impl\n    Set b = a\n    Set c = a\n    Set o = a\nEnd Sub\n"
    )
    found = _project(
        ("IFoo", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Sub Run()\nEnd Sub\n"),
        ("IBar", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Sub Halt()\nEnd Sub\n"),
        ("Impl", ModuleSymbolKind.CLASS, impl),
        ("Other", ModuleSymbolKind.CLASS, "Option Explicit\nPublic Sub Walk()\nEnd Sub\n"),
        ("Module1", ModuleSymbolKind.STANDARD, main),
    )
    assert [f for f in found["Module1"] if f[0] == "assignment-object-type-mismatch"] == [
        (
            "assignment-object-type-mismatch",
            "a",
            "Object assignment to 'o' expects Other, but got a As IFoo. This object type is not "
            "compatible with Other.",
        )
    ]


def test_an_object_declared_item_is_late_bound() -> None:
    def member_not_found(line: str) -> list[tuple[str, str]]:
        return _found(f"Option Explicit\nSub S()\n    {line}\nEnd Sub\n", "member-not-found")

    assert member_not_found("Worksheets(1).NoSuchMemberXyz") == []
    assert member_not_found("ActiveSheet.NoSuchMemberXyz") == []
    assert member_not_found("Dim ws As Worksheet: ws.NoSuchMemberXyz") == [
        ("NoSuchMemberXyz", "Method or data member not found: 'Excel.Worksheet.NoSuchMemberXyz'.")
    ]


def test_word_powerpoint_and_access_types_carry_their_measured_flags() -> None:
    assert host_type_resolves_when_compiling("Word.Range")
    assert not host_type_resolves_when_compiling("Word.Document")
    assert host_type_resolves_when_compiling("PowerPoint.Slide")
    assert host_type_resolves_when_compiling("Access.TextBox")
    assert not host_type_resolves_when_compiling("Access.Form")
    assert not host_type_resolves_when_compiling("Office.CommandBar")


# -- arguments, assignments and declarations -------------------------------------


def test_a_variant_variable_passed_byref_to_a_typed_parameter() -> None:
    source = (
        "Option Explicit\nSub Take(x As Long)\nEnd Sub\nSub S()\n    Const K As Integer = 1\n"
        "    Dim v As Variant\n    Take v\n    Take (v)\n    Take K\nEnd Sub\n"
    )
    assert _found(source, "byref-argument-type-mismatch") == [
        (
            "v",
            "ByRef argument 'x' of 'Take' expects Long, but 'v' is declared as Variant. This is a VBE "
            "compile error: ByRef argument type mismatch.",
        )
    ]


def test_cstr_of_null_raises_where_left_hands_null_back() -> None:
    source = "Option Explicit\nSub S()\n    Dim s As String\n    s = CStr(Null)\n    s = Left(Null, 1) & \"\"\nEnd Sub\n"
    assert _found(source, "argument-type-mismatch") == [
        ("Null", "Argument 'Expression' of 'CStr' cannot be Null. This will raise Run-time error '94': Invalid use of Null.")
    ]


def test_statement_forms_that_assign_a_function_result() -> None:
    source = (
        "Option Explicit\nFunction Count3() As Long\n    For Count3 = 1 To 3\n    Next\nEnd Function\n"
        "Function Three() As Variant\n    ReDim Three(2)\nEnd Function\n"
        "Function Named() As String\n    Named$ = \"hi\"\nEnd Function\n"
    )
    assert "missing-return-assignment" not in _codes(source)


def test_a_dynamic_byte_array_takes_a_string() -> None:
    source = "Option Explicit\nSub S()\n    Dim b() As Byte\n    b = \"abc\"\n    Debug.Print b(0)\nEnd Sub\n"
    assert _codes(source) == set()


def test_set_of_a_literal_is_refused() -> None:
    source = "Option Explicit\nSub S()\n    Dim v As Variant\n    Set v = 5\n    Debug.Print v\nEnd Sub\n"
    assert _found(source, "set-requires-object") == [
        ("5", "Set assigns an object reference, but 5 is a literal value. This is a VBE compile error: Object required.")
    ]


def test_duplicate_names_the_vbe_refuses() -> None:
    assert _found("Option Explicit\nFunction Main() As Long\n    Dim Main As Long\nEnd Function\n", "duplicate-declaration") == [
        ("Main", "Duplicate declaration in current scope: 'Main'.")
    ]
    clash = "Option Explicit\nPrivate Helper As Long\nPrivate Sub Helper()\nEnd Sub\n"
    assert _found(clash, "duplicate-procedure") == [
        ("Helper", "Ambiguous name detected: 'Helper' names both a variable and a procedure in this module.")
    ]


def test_property_procedures_must_agree_on_the_value_type() -> None:
    source = (
        "Option Explicit\nPublic Property Get Size() As Long\n    Size = 1\nEnd Property\n"
        "Public Property Let Size(ByVal v As Integer)\nEnd Property\n"
    )
    assert _found(source, "property-accessor-signature-mismatch") == [
        (
            "v",
            "Property Let 'Size' takes its value As Integer, but Property Get 'Size' returns Long; the "
            "definitions of a property's procedures must agree.",
        )
    ]


def test_withevents_needs_a_class_that_sources_events() -> None:
    source = (
        "Option Explicit\nPrivate WithEvents a As Object\nPrivate WithEvents b As Collection\n"
        "Private WithEvents c As Long\nPrivate WithEvents d As Worksheet\n"
    )
    found = _found(source, "withevents-declaration", AnalyzeModuleOptions(module_kind=ModuleSymbolKind.CLASS))
    assert found == [
        ("a", "WithEvents variable 'a' must be declared As a specific class that raises events; 'Object' names none."),
        ("b", "WithEvents variable 'b' is declared As Collection, which does not source automation events."),
        ("c", "WithEvents variable 'c' is declared As Long, which does not source automation events."),
    ]


def test_the_vbe_folds_some_intrinsics_in_an_enum_but_none_in_a_const() -> None:
    source = (
        "Option Explicit\nPrivate Enum E\n    eA = Len(\"ab\")\n    eB = Asc(\"a\")\nEnd Enum\n"
        "Private Const C As Long = Len(\"ab\")\nSub S(Optional ByVal n As Long = CLng(2))\n    Debug.Print C\nEnd Sub\n"
    )
    assert _found(source, "enum-member-not-constant") == [
        ("Asc(\"a\")", "Enum member 'eB' value must be a constant expression; the call 'Asc(...)' is not constant.")
    ]
    assert _found(source, "const-value-not-constant") == [
        ("Len(\"ab\")", "Const 'C' value must be a constant expression; the call 'Len(...)' is not constant.")
    ]
    assert "parameter-default-not-constant" not in _codes(source)


def test_only_a_procedure_closes_the_option_window() -> None:
    assert "option-after-declaration" not in _codes("Private Const A As Long = 1\nOption Explicit\n")
    assert _found("Option Explicit\nPrivate arr(3) As Long\nOption Base 1\n", "option-after-declaration") == [
        ("Option", "'Option Base' must come before any array declaration: an array above it is already dimensioned.")
    ]
    private = "Option Explicit\nOption Private Module\n"
    assert _found(private, "invalid-option-statement", AnalyzeModuleOptions(module_kind=ModuleSymbolKind.CLASS)) == [
        ("Module", "'Option Private Module' is not permitted in a class, document or UserForm module.")
    ]


# -- control flow, names and expressions -----------------------------------------


def test_open_needs_no_for_but_a_mode_word_needs_its_for() -> None:
    assert _found("Sub S\n    Open \"f.txt\" Output As #1\nEnd Sub", "open-missing-for") == [
        ("Open", "An 'Open' statement's mode 'Output' must follow 'For'.")
    ]
    assert "open-missing-for" not in _codes("Sub S\n    Open \"f.txt\" As #1\nEnd Sub")


def test_for_each_over_an_array_needs_a_variant() -> None:
    source = "Option Explicit\nSub S()\n    Dim arr(1) As Variant, o As Object\n    For Each o In arr\n    Next\nEnd Sub\n"
    assert _found(source, "for-each-control-variable-type") == [
        ("o", "For Each control variable 'o' must be Variant when the source is an array, but it is declared As Object.")
    ]


def test_a_module_name_alone_is_not_a_call() -> None:
    found = _project(
        ("Helpers", ModuleSymbolKind.STANDARD, "Option Explicit\nPublic Sub Bar()\nEnd Sub\n"),
        ("Module1", ModuleSymbolKind.STANDARD, "Option Explicit\nPublic Sub Main()\n    Helpers\n    Helpers.Bar\nEnd Sub\n"),
    )
    assert [f for f in found["Module1"] if f[0] == "unknown-call"] == [
        (
            "unknown-call",
            "Helpers",
            "'Helpers' is a module, not a procedure: name the procedure to call, as in 'Helpers.Bar'. "
            "This is a VBE compile error: Expected variable or procedure, not module.",
        )
    ]


def test_redim_declares_and_a_library_name_qualifies() -> None:
    assert _codes("Option Explicit\nSub S()\n    ReDim items(2) As Long\n    items(0) = 1\nEnd Sub\n") == set()
    qualified = "Option Explicit\nSub S()\n    Dim app As Object\n    Set app = Excel.Application\n    Debug.Print app.Name\nEnd Sub\n"
    assert _codes(qualified) == set()


def test_object_state_follows_for_each_and_one_line_guards() -> None:
    source = (
        "Option Explicit\nSub S()\n    Dim ws As Worksheet, c As Collection, d As Collection\n"
        "    For Each ws In c\n    Next\n    If Not d Is Nothing Then d.Add 1\nEnd Sub\n"
    )
    assert _found(source, "object-variable-not-set") == [
        ("c", "Object variable 'c' is Nothing when For Each asks it for its elements. This will raise Run-time error '424': Object required.")
    ]


def test_division_by_zero_follows_guards_known_locals_and_rounding() -> None:
    source = (
        "Option Explicit\nPrivate Const SCALE_BY As Long = 0\nSub S()\n    Dim x As Double, d As Long\n"
        "    If SCALE_BY <> 0 Then x = 10 / SCALE_BY\n    x = 10 / d\n    x = 0 / 0\n    x = 5 \\ 0.4\n"
        "    Debug.Print x\nEnd Sub\n"
    )
    assert _found(source, "division-by-zero") == [
        ("d", "Expression uses '/' with a zero divisor. This will raise Run-time error '11': Division by zero."),
        ("0", "Expression divides zero by zero with '/'. This will raise Run-time error '6': Overflow."),
        ("0.4", "Expression uses '\\' with a zero divisor. This will raise Run-time error '11': Division by zero."),
    ]


def test_a_nonnumeric_string_operand_raises_whatever_the_target() -> None:
    source = (
        "Option Explicit\nFunction Main() As Variant\n    Dim v As Variant, s As String\n    s = \"abc\"\n"
        "    v = \"abc\" + 1\n    v = s * 2\n    Main = v\nEnd Function\n"
    )
    assert _found(source, "string-arithmetic-coercion") == [
        ("\"abc\"", "Operator '+' coerces string literal \"abc\" to a number. This will raise Run-time error '13': Type mismatch."),
        ("s", "Operator '*' coerces 's', which holds \"abc\" to a number. This will raise Run-time error '13': Type mismatch."),
    ]


def test_a_glued_ampersand_is_a_type_suffix() -> None:
    assert _codes("Option Explicit\nSub S()\n    Dim total&\n    total& = 3\n    Debug.Print total\nEnd Sub\n") == set()


def test_literal_forms_the_vbe_refuses() -> None:
    source = (
        "Option Explicit\nSub S()\n    Dim v As Variant\n    v = 2147483648&\n    v = &H\n    v = 3.5E+38!\n"
        "    v = 922337203685477.5808@\n    v = #2/30/2000#\n    v = #25:00#\n    Debug.Print v\nEnd Sub\n"
    )
    assert _found(source, "suffixed-literal-overflow") == [
        (
            "2147483648&",
            "The literal '2147483648&' is outside the Long range -2147483648 to 2147483647 of its '&' "
            "type suffix. VBE rejects this at compile time as a Syntax error.",
        ),
        ("&H", "'&H' names a radix with no digits after it. VBE rejects this at compile time as a Syntax error."),
    ]
    assert [text for text, _ in _found(source, "float-literal-overflow")] == ["3.5E+38!", "922337203685477.5808@"]
    assert _found(source, "date-literal-invalid") == [
        ("#2/30/2000#", "The date literal #2/30/2000# names day 30 in a month of 29 days. VBE rejects this at compile time as a Syntax error."),
        ("#25:00#", "The date literal #25:00# names hour 25; hours run 0 to 23. VBE rejects this at compile time as a Syntax error."),
    ]


# -- new rules -------------------------------------------------------------------


def _main(*lines: str) -> str:
    body = "".join(f"    {line}\n" for line in lines)
    return f"Option Explicit\nFunction Main() As Variant\n{body}End Function\n"


def test_overflow_the_analyzer_can_prove() -> None:
    assert _found(_main("Dim secs As Long", "secs = 60 * 60 * 24", "Main = secs"), "arithmetic-overflow") == [
        (
            "60 * 60 * 24",
            "3600 (Integer) * 24 (Integer) is 86400, outside the Integer range. This will raise "
            "Run-time error '6': Overflow.",
        )
    ]
    assert _found(_main("Dim b As Byte", "b = 255.5", "Main = b"), "arithmetic-overflow") == [
        (
            "255.5",
            "Assignment to 'b' stores 256 (255.5 rounds to 256) in a Byte, whose range is 0 to 255. "
            "This will raise Run-time error '6': Overflow.",
        )
    ]
    constant = "Option Explicit\nPrivate Const SECONDS_PER_DAY = 60 * 60 * 24\nSub S()\n    Debug.Print SECONDS_PER_DAY\nEnd Sub\n"
    assert _found(constant, "const-overflow") == [
        (
            "60 * 60 * 24",
            "Const 'SECONDS_PER_DAY' overflows while it is evaluated: 3600 (Integer) * 24 (Integer) is "
            "86400, outside the Integer range. This is a VBE compile error: Overflow.",
        )
    ]
    loop = _main("Dim i As Integer", "For i = 1 To 32767", "Next", "Main = i")
    assert _found(loop, "for-counter-overflow") == [
        (
            "32767",
            "Counter 'i' is Integer; after its last pass at 32767 the loop adds 1, which does not fit. "
            "This will raise Run-time error '6': Overflow.",
        )
    ]
    assert "for-counter-overflow" not in _codes(_main("Dim i As Integer", "For i = 1 To 32766", "Next", "Main = i"))


def test_collection_contents_the_code_makes_plain() -> None:
    assert _found(_main("Dim c As New Collection", "c.Add 1", "c.Add 2", "Main = c(3)"), "collection-index-out-of-range") == [
        (
            "3",
            "'c' holds 2 elements here, indexed 1 to 2; 3 is outside that. This will raise Run-time "
            "error '9': Subscript out of range.",
        )
    ]
    assert _found(_main("Dim c As New Collection", "c.Add 1, \"k\"", "c.Add 2, \"K\""), "collection-key-in-use") == [
        (
            "\"K\"",
            "'c' already has an element with the key \"K\" (keys compare without case). This will raise "
            "Run-time error '457': This key is already associated with an element of this collection.",
        )
    ]
    quiet = _main("Dim c As New Collection", "c.Add 1, \"k\"", "Main = c(\"K\")", "If Main Then c.Add 3", "Main = c(3)")
    assert not _codes(quiet) & {"collection-index-out-of-range", "collection-key-not-found"}


def test_file_statements_whose_failure_the_code_proves() -> None:
    target = "Environ$(\"TEMP\") & \"\\a.txt\""
    assert _found(_main(f"Open {target} For Output As #7", f"Open {target} For Output As #7"), "file-already-open") == [
        (
            "7",
            "File number #7 is still open from the Open statement above; opening it again raises "
            "Run-time error '55': File already open. Close it first.",
        )
    ]
    closed = _main("Dim f As Integer", "f = FreeFile", f"Open {target} For Output As #f", "Close #f", "Print #f, \"late\"")
    assert _found(closed, "file-used-after-close") == [
        ("f", "File number 'f' was closed above and not opened again. This will raise Run-time error '52': Bad file name or number.")
    ]
    assert _found(_main("Main = LOF(0)"), "file-number-zero") == [
        ("0", "File number 0 is never open: file numbers run from 1 to 511. This will raise Run-time error '52': Bad file name or number.")
    ]


def test_host_arguments_the_code_proves_wrong() -> None:
    assert _found(_main("Main = Worksheets(0).Name"), "host-argument-out-of-range") == [
        ("0", "Index 0 is never an element: Excel collections start at 1. This will raise Run-time error '9': Subscript out of range.")
    ]
    assert _found(_main("Worksheets(1).Name = \"a:b\""), "sheet-name-invalid") == [
        (
            "\"a:b\"",
            "Excel refuses this name: a sheet name cannot contain any of : \\ / ? * [ ]. This will raise "
            "Run-time error '1004': You typed an invalid name for a sheet or chart.",
        )
    ]
    assert _found(_main("Dim s As String", "s = Range(\"A1:B2\")", "Main = s"), "multi-cell-range-as-scalar") == [
        (
            "Range(\"A1:B2\")",
            "Range(\"A1:B2\") read as a value is a two-dimensional array, which a String variable cannot "
            "hold. This will raise Run-time error '13': Type mismatch.",
        )
    ]


def test_a_late_bound_member_that_fails_when_it_runs() -> None:
    assert _found(_main("Application.Zzq"), "runtime-member-not-found") == [
        (
            "Zzq",
            "Application has no member 'Zzq', and it is not a worksheet function either. The VBE compiles "
            "the name because Application is extensible; this will raise Run-time error '438': Object "
            "doesn't support this property or method.",
        )
    ]
    assert "runtime-member-not-found" not in _codes(_main("Main = Application.Match(1, Array(1), 0)"))


def test_declaration_forms_the_vbe_refuses() -> None:
    source = (
        "Option Explicit\nPrivate Sub F(ByVal a() As Long, Optional ByVal i As Integer = 40000)\nEnd Sub\n"
        "Private Sub G()\n    Dim [my var] As Long\nEnd Sub\n"
    )
    assert _found(source, "array-parameter-form") == [
        ("a", "Array parameter 'a' cannot be ByVal: an array argument must be ByRef.")
    ]
    assert _found(source, "parameter-default-type-mismatch") == [
        (
            "40000",
            "Optional parameter 'i' is declared As Integer, whose range is -32768 to 32767; its default "
            "40000 does not fit. This is a VBE compile error: Overflow.",
        )
    ]
    assert _found(source, "bracketed-variable-name") == [
        (
            "[my var]",
            "'[my var]' is not a variable name: brackets make a foreign name only for an Enum member or a "
            "member of another object. This is a VBE compile error: Syntax error.",
        )
    ]
    assert _found("Option Explicit\nDefLng A-Z\nDefStr S\n", "duplicate-deftype") == [
        (
            "DefStr",
            "Letter 'S' already has a default type from 'DefLng' above; a letter takes one Deftype "
            "statement. This is a VBE compile error: Duplicate Deftype statement.",
        )
    ]


def test_line_continuation_limits() -> None:
    blank = "Option Explicit\nSub S()\n    Dim x As Long\n    x = 1 + _\n\n    Debug.Print x\nEnd Sub\n"
    assert _found(blank, "invalid-line-continuation")[0][1] == (
        "A line continuation must be followed by more of the statement; the next line is empty. This "
        "is a VBE compile error: Syntax error."
    )
    enum = "Option Explicit\nPrivate Enum K\n    kOne _\n    = 1\nEnd Enum\n"
    assert _found(enum, "invalid-line-continuation") == [
        (
            " _\n",
            "A line continuation is not allowed inside Enum 'K': neither on a member line nor before End "
            "Enum (\"Invalid inside Enum\").",
        )
    ]


def test_directive_forms_the_vbe_refuses() -> None:
    assert _found("Option Explicit\n#Const FEATURE = 1\n#Const FEATURE = 2\n", "duplicate-const-directive") == [
        ("FEATURE", "'#Const FEATURE' is already defined in this module. This is a VBE compile error: Duplicate definition.")
    ]
    trailing = "Option Explicit\nSub S()\n#If VBA7 Then: Debug.Print 1\n#End If\nEnd Sub\n"
    assert [text for text, _ in _found(trailing, "directive-trailing-statement")] == ["Debug.Print 1"]


def test_stray_characters_and_long_lines() -> None:
    assert _found("Option Explicit\nSub S()\n    Dim n As Long\n    n = 1;\n    Debug.Print n; n\nEnd Sub\n", "stray-character") == [
        (";", "A ';' ends no VBA statement: it separates items only in a Print or Write list. Remove it. This is a VBE compile error: Syntax error.")
    ]
    long_line = "Option Explicit\nSub S()\n    Debug.Print \"" + "x" * 1100 + "\"\nEnd Sub\n"
    assert [message for _, message in _found(long_line, "line-too-long")] == [
        "Line 3 is 1118 characters long; the VBE accepts 1023. Break it with a line continuation. This is a VBE compile error."
    ]


def test_handler_flow_errors() -> None:
    falls_in = _main("On Error GoTo Handler", "Main = 1", "Handler:", "Err.Raise Err.Number")
    assert _found(falls_in, "handler-fall-through") == [
        (
            "Handler",
            "Execution falls into error handler 'Handler' with no error pending, and 'Err.Raise "
            "Err.Number' then raises with Err.Number 0. This will raise Run-time error '5': Invalid "
            "procedure call or argument. Put an Exit Function before the label.",
        )
    ]
    assert "handler-fall-through" not in _codes(
        _main("On Error GoTo Handler", "Main = 1", "Exit Function", "Handler:", "Err.Raise Err.Number")
    )
    assert _found(_main("Main = 1", "Resume Next"), "resume-without-error") == [
        (
            "Resume",
            "'Resume' runs with no error handler installed in this procedure, so no error is pending. "
            "This will raise Run-time error '20': Resume without error.",
        )
    ]
    gosub = _main("GoSub Work", "Work:", "Main = Main + 1", "Return")
    assert [text for text, _ in _found(gosub, "return-without-gosub")] == ["Work"]


def test_a_property_accessor_that_calls_itself() -> None:
    getter = "Option Explicit\nPublic Property Get Name() As String\n    Name = Me.Name\nEnd Property\n"
    assert _found(getter, "recursive-property-accessor", AnalyzeModuleOptions(module_kind=ModuleSymbolKind.CLASS)) == [
        (
            "Me.Name",
            "Property Get 'Name' reads 'Me.Name', which is itself: the call never returns. This will "
            "raise Run-time error '28': Out of stack space.",
        )
    ]


def test_statement_forms_the_vbe_refuses() -> None:
    assert _found(_main("Dim x As Boolean", "If x Then Rem note"), "rem-after-then") == [
        (
            "Rem",
            "'Rem' cannot follow 'Then' on one line: a Rem comment starts only at the start of a "
            "statement. This is a VBE compile error: Syntax error.",
        )
    ]
    operand = _main("Dim c As New Collection, x As Variant", "x = c + 1", "Main = x")
    assert _found(operand, "collection-operand") == [
        (
            "c",
            "'c' is a Collection: its default member Item needs an index, so '+' has no value to work "
            "on. This is a VBE compile error: Argument not optional.",
        )
    ]
    sub_value = (
        "Option Explicit\nPrivate Sub Foo()\nEnd Sub\nPrivate Function Bar() As Long\nEnd Function\n"
        "Function Main() As Variant\n    Dim x As Variant\n    x = Foo\n    x = Bar\n    Foo\n"
        "    Main = x\nEnd Function\n"
    )
    assert _found(sub_value, "sub-used-as-value") == [
        (
            "Foo",
            "'Foo' is a Sub, which returns nothing, so it cannot be used as a value. This is a VBE "
            "compile error: Expected Function or variable.",
        )
    ]


def test_implements_needs_every_member_with_its_signature() -> None:
    interface = "Option Explicit\nPublic Function Name() As String\nEnd Function\nPublic Sub Size(ByVal n As Long)\nEnd Sub\n"
    implementer = (
        "Option Explicit\nImplements IFoo\nPrivate Function IFoo_Name() As String\nEnd Function\n"
        "Private Sub IFoo_Size(ByVal n As Integer)\nEnd Sub\n"
    )
    found = _project(("IFoo", ModuleSymbolKind.CLASS, interface), ("CFoo", ModuleSymbolKind.CLASS, implementer))
    assert [f for f in found["CFoo"] if f[0].startswith("implements-member")] == [
        (
            "implements-member-signature",
            "IFoo_Size",
            "'IFoo_Size' does not match 'IFoo.Size': parameter 1 is Integer here and long on the "
            "interface. The procedure declaration must match the interface member it implements.",
        )
    ]


def test_runtime_values_the_code_states() -> None:
    assert _found(_main("Main = Asc(\"\")"), "runtime-argument-value") == [
        ("\"\"", "Argument 'String' of 'Asc' is \"\"; this will raise Run-time error '5': Invalid procedure call or argument.")
    ]
    mid = _main("Dim s As String", "s = \"abc\"", "Mid(s, 5, 1) = \"x\"", "Main = s")
    assert _found(mid, "runtime-argument-value") == [
        (
            "5",
            "Mid statement start 5 is past the end of s, which is 3 character(s) long. This will raise "
            "Run-time error '5': Invalid procedure call or argument.",
        )
    ]
    assert [text for text, _ in _found(_main("Err.Raise 0", "Err.Raise 5"), "runtime-argument-value")] == ["0"]
    assert _found(_main("Main = \"b\" Like \"[z-a]\""), "runtime-argument-value") == [
        ("\"[z-a]\"", "The Like pattern \"[z-a]\" has the reversed range z-a. This will raise Run-time error '93': Invalid pattern string.")
    ]
    assert _found(_main("Main = CBool(\"yes\")"), "runtime-conversion-value") == [
        ("\"yes\"", "CBool cannot convert \"yes\" to Boolean. This will raise Run-time error '13': Type mismatch.")
    ]


def test_array_bounds_the_code_makes_plain() -> None:
    split = _main("Dim a() As String", "a = Split(\"a,b\", \",\")", "Main = a(2)")
    assert _found(split, "array-subscript-out-of-bounds") == [
        (
            "2",
            "Subscript 2 for array 'a' (Split(...)) is above the upper bound 1. This will raise "
            "Run-time error '9': Subscript out of range.",
        )
    ]
    two_d = _main("Dim a(1 To 3, 1 To 2) As Long", "Main = a(2, 3)")
    assert [message for _, message in _found(two_d, "array-subscript-out-of-bounds")] == [
        "Subscript 3 for array 'a' is above the upper bound 2 in dimension 2. This will raise "
        "Run-time error '9': Subscript out of range."
    ]
    loop = _main("Dim a(2) As Long, i As Long", "For i = 0 To 3", "a(i) = i", "Next", "Main = a(0)")
    assert [message for _, message in _found(loop, "array-subscript-out-of-bounds")] == [
        "Counter 'i' reaches 3 on its last pass, which for array 'a' is above the upper bound 2. "
        "This will raise Run-time error '9': Subscript out of range."
    ]


def test_a_variant_used_as_the_wrong_kind_of_value() -> None:
    assert _found(_main("Dim v As Variant", "v = 5", "v.Foo"), "variant-value-misuse") == [
        ("v", "'v' holds the number 5 here, which has no members. This will raise Run-time error '424': Object required.")
    ]
    array = _main("Dim v As Variant", "v = Array(1, 2)", "Main = v + 1")
    assert _found(array, "variant-value-misuse") == [
        (
            "v",
            "'v' holds an array from Array(...) here, which '+' cannot combine with a scalar. This "
            "will raise Run-time error '13': Type mismatch.",
        )
    ]
    quiet = _main("Dim v As Variant", "v = Array(1, 2)", "Main = v(0) + 1", "Main = UBound(v)")
    assert "variant-value-misuse" not in _codes(quiet)
