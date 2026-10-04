"""handlerFlow and undeclared behaviour added by the sync to XLIDE 2f49b93 (10.14.6).

Each expectation was run through the pinned upstream analyzer
(artifacts/analyzer-pin/2f49b93, tools/hunt/wrapper_probe.mjs) with the same
findings.
"""

from __future__ import annotations

from pyvbaanalysis import analyze_project
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_STD = ModuleSymbolKind.STANDARD
_CLASS = ModuleSymbolKind.CLASS


def _found(*modules: tuple[str, ModuleSymbolKind, str]) -> dict[str, list[tuple[str, str, str]]]:
    """(code, covered text, message) per module."""
    sources = {name: source for name, _, source in modules}
    results = analyze_project([ModuleInput(name, kind, source) for name, kind, source in modules])
    return {
        name: [(d.code, sources[name][d.span.start : d.span.end], d.message) for d in found]
        for name, found in results.items()
    }


def _codes(code: str, *modules: tuple[str, ModuleSymbolKind, str]) -> list[tuple[str, str]]:
    return [
        (text, message)
        for found in _found(*modules).values()
        for one, text, message in found
        if one == code
    ]


_STACK = "This will raise Run-time error '28': Out of stack space."


# -- #240 and #613: unbounded recursion --------------------------------------------


def test_a_procedure_that_calls_itself_first_never_returns() -> None:
    source = "Option Explicit\r\nSub S()\r\n    S\r\nEnd Sub\r\n"
    assert _codes("unbounded-recursion", ("Module1", _STD, source)) == [
        ("S", f"'S' calls itself before anything could make it return: the calls never end. {_STACK}")
    ]


def test_a_cycle_of_first_calls_reports_each_procedure() -> None:
    source = (
        "Option Explicit\r\nSub A()\r\n    B\r\nEnd Sub\r\nSub B()\r\n    C\r\nEnd Sub\r\n"
        "Sub C()\r\n    A\r\nEnd Sub\r\n"
    )
    messages = [message for _, message in _codes("unbounded-recursion", ("Module1", _STD, source))]
    assert messages == [
        f"'A' calls 'B', which calls 'C', which calls 'A', before anything could make it return: "
        f"the calls never end. {_STACK}",
        f"'B' calls 'C', which calls 'A', which calls 'B', before anything could make it return: "
        f"the calls never end. {_STACK}",
        f"'C' calls 'A', which calls 'B', which calls 'C', before anything could make it return: "
        f"the calls never end. {_STACK}",
    ]


def test_a_call_inside_a_single_line_if_is_conditional() -> None:
    source = "Option Explicit\r\nSub S(n)\r\n    If n > 0 Then S n - 1\r\nEnd Sub\r\n"
    assert _codes("unbounded-recursion", ("Module1", _STD, source)) == []


def test_call_by_name_on_me_calls_the_method() -> None:
    source = 'Option Explicit\r\nPublic Sub Go()\r\n    CallByName Me, "Go", VbMethod\r\nEnd Sub\r\n'
    assert [text for text, _ in _codes("unbounded-recursion", ("Class1", _CLASS, source))] == [
        'CallByName Me, "Go"'
    ]


# -- #338 and #613: recursive property accessors ----------------------------------


def test_a_property_reads_itself_through_with_me_and_a_new_instance() -> None:
    with_me = (
        "Option Explicit\r\nProperty Get Value() As Long\r\n    With Me\r\n"
        "        Value = .Value\r\n    End With\r\nEnd Property\r\n"
    )
    assert [text for text, _ in _codes("recursive-property-accessor", ("Class1", _CLASS, with_me))] == [
        ".Value"
    ]
    alias = (
        "Option Explicit\r\nProperty Get Value() As Long\r\n    Dim o As Class1\r\n"
        "    Set o = New Class1\r\n    Value = o.Value\r\nEnd Property\r\n"
    )
    assert [text for text, _ in _codes("recursive-property-accessor", ("Class1", _CLASS, alias))] == [
        "o.Value"
    ]


def test_a_property_let_assigns_itself_through_me() -> None:
    source = "Option Explicit\r\nProperty Let Value(v As Long)\r\n    Me.Value = v\r\nEnd Property\r\n"
    assert _codes("recursive-property-accessor", ("Class1", _CLASS, source)) == [
        (
            "Me.Value",
            f"Property Let 'Value' assigns 'Me.Value', which is itself: the call never returns. {_STACK}",
        )
    ]


# -- #237 and #203: labels in blocks, and dead code above a label ------------------


def test_a_handler_label_inside_a_block_is_fallen_into() -> None:
    source = (
        "Option Explicit\r\nSub S()\r\n    Dim x\r\n    On Error GoTo H\r\n    If True Then\r\n"
        "        x = 1\r\nH:\r\n        Err.Raise Err.Number\r\n    End If\r\nEnd Sub\r\n"
    )
    assert [text for text, _ in _codes("handler-fall-through", ("Module1", _STD, source))] == ["H"]


def test_dead_code_above_a_gosub_target_falls_into_nothing() -> None:
    source = (
        "Option Explicit\r\nSub S()\r\n    Dim x\r\n    GoSub T\r\n    Exit Sub\r\nSkip:\r\n"
        "    x = 1\r\nT:\r\n    Return\r\nEnd Sub\r\n"
    )
    assert _codes("return-without-gosub", ("Module1", _STD, source)) == []


# -- #318, #369, #445, #266, #224: undeclared names ---------------------------------


def test_a_library_procedure_named_bare_as_a_value() -> None:
    source = (
        "Option Explicit\r\nSub Main()\r\n    Dim s\r\n    s = Beep\r\n"
        "    Debug.Print TypeName(Kill)\r\nEnd Sub\r\n"
    )
    found = _found(("Module1", _STD, source))["Module1"]
    assert ("sub-used-as-value", "Beep") in [(code, text) for code, text, _ in found]
    assert ("argument-count", "Kill") in [(code, text) for code, text, _ in found]


def test_names_in_const_enum_and_optional_default_values() -> None:
    source = (
        "Option Explicit\r\nConst K = asdf\r\nEnum E\r\n    eA\r\n    eB = qwer\r\nEnd Enum\r\n"
        "Sub S(Optional x As Long = y)\r\n    Const L = zz + 1\r\n    Debug.Print K, L, x\r\nEnd Sub\r\n"
    )
    assert _codes("undeclared-variable", ("Module1", _STD, source)) == [
        ("asdf", "Variable not defined: 'asdf'. Declare it before using it in a Const's value, or remove Option Explicit."),
        (
            "qwer",
            "'qwer' is not defined, and an Enum member's value must be a constant. "
            "This is a VBE compile error: Constant expression required.",
        ),
        (
            "y",
            "Variable not defined: 'y'. Declare it before using it in an Optional parameter's default, "
            "or remove Option Explicit.",
        ),
        ("zz", "Variable not defined: 'zz'. Declare it before using it in a Const's value, or remove Option Explicit."),
    ]


def test_an_event_alone_is_no_procedure_to_call() -> None:
    source = "Option Explicit\r\nEvent Done()\r\nSub Go()\r\n    Done\r\nEnd Sub\r\n"
    assert _codes("unknown-call", ("Class1", _CLASS, source)) == [
        ("Done", "Sub or Function not defined: 'Done'.")
    ]


def test_class_members_in_forms_the_vbe_refuses() -> None:
    class1 = (
        "Option Explicit\r\nPublic Field As Long\r\nPublic Property Get Idx(i As Long) As Long\r\n"
        "End Property\r\n"
    )
    module1 = (
        "Option Explicit\r\nSub Main()\r\n    Dim c As New Class1\r\n    Debug.Print c.Field(1)\r\n"
        "    c.Field(1) = 5\r\n    Debug.Print c.Idx\r\nEnd Sub\r\n"
    )
    found = _found(("Class1", _CLASS, class1), ("Module1", _STD, module1))["Module1"]
    assert [(text, message) for code, text, message in found if code == "argument-count"] == [
        (
            "Field",
            "'Field' is a field of type Long, which takes no arguments. This is a VBE compile error: "
            "Wrong number of arguments or invalid property assignment.",
        ),
        (
            "Field",
            "'Field' is a field of type Long, which takes no index, so 'Field(...)' is no place to "
            "assign. This is a VBE compile error: Can't assign to read-only property.",
        ),
        (
            "Idx",
            "Argument not optional: property 'Idx' takes an index, as in Idx(i As Long) As Long. "
            "This is a VBE compile error.",
        ),
    ]
