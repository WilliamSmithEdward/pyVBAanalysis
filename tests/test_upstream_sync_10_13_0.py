"""Behaviour changed by the sync to XLIDE 10.13.0 (issue #139).

10.13.0 is upstream's fix for its 10.12.0 slowdown, and two findings changed with
it; both expectations were run through the pinned upstream analyzer as well. The
lexing guard itself is test_analyzer_lexes_module_once.py.

* An Enum member value or an Optional default is read as the expression it is: its
  leading number is no longer dropped as if it were a line label, so
  `= 300 - 100` is not the literal -100.
* The module's Consts are folded once and a procedure's own Consts layer on top, so
  a module Const keeps the value its own module gives it when a procedure declares a
  Const of the same name as one it depends on.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics import analyze_module


def _found(source: str, code: str) -> list[tuple[str, str]]:
    return [(source[d.span.start : d.span.end], d.message) for d in analyze_module(source) if d.code == code]


def test_a_default_or_enum_value_keeps_its_leading_number() -> None:
    default = (
        "Option Explicit\nPrivate Sub F(Optional ByVal b As Byte = 300 - 100, "
        "Optional ByVal i As Integer = 40000)\nEnd Sub\n"
    )
    assert [text for text, _ in _found(default, "parameter-default-type-mismatch")] == ["40000"]
    enum = "Option Explicit\nPrivate Enum E\n    eH = 1 - 3000000000\n    eX = 3000000000#\nEnd Enum\n"
    assert [text for text, _ in _found(enum, "const-overflow")] == ["eX"]


def test_a_module_const_keeps_its_own_module_value_inside_a_shadowing_procedure() -> None:
    source = (
        "Option Explicit\nPrivate Const B As Integer = 2\nPrivate Const A As Integer = B * 10000\n"
        "Sub S()\n    Const B As Integer = 4\n    Dim y As Integer\n    y = A * 2\n    Debug.Print y + B\nEnd Sub\n"
    )
    assert _found(source, "arithmetic-overflow") == [
        (
            "A * 2",
            "20000 (Integer) * 2 (Integer) is 40000, outside the Integer range. This will raise "
            "Run-time error '6': Overflow.",
        )
    ]
