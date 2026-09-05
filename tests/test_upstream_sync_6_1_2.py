"""Behaviour added by the sync to XLIDE 6.1.2.

Two changes with teeth:

* ``ambiguousProjectProcedure`` (new rule): VBA refuses to compile an unqualified
  call to a name two modules export.
* ``missingReturnAssignment`` widened from untyped functions only to every
  Function and Property Get, which required three detector fixes so the widening
  did not manufacture false positives.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_STD = ModuleSymbolKind.STANDARD
_CODE = "ambiguous-project-procedure"


def _codes(source: str, **kwargs: object) -> list[str]:
    return [
        d.code
        for d in analyze_module(source, AnalyzeModuleOptions(**kwargs))  # type: ignore[arg-type]
        if d.code != "option-explicit-missing"
    ]


def _project_codes(modules: list[ModuleInput], name: str) -> list[str]:
    return [d.code for d in analyze_project(modules)[name] if d.code != "option-explicit-missing"]


def _caller_codes(caller: str, *, second_exporter: str | None = None) -> list[str]:
    modules = [ModuleInput("Helpers", _STD, "Public Sub Recalculate()\nEnd Sub\n")]
    if second_exporter is not None:
        modules.append(ModuleInput("Utils", _STD, second_exporter))
    modules.append(ModuleInput("Caller", _STD, caller))
    return [d.code for d in analyze_project(modules)["Caller"]]


# -- ambiguous unqualified call --------------------------------------------

_EXPORTS = "Public Sub Recalculate()\nEnd Sub\n"


def test_a_bare_call_to_a_name_two_modules_export_reports() -> None:
    codes = _caller_codes(
        "Option Explicit\nSub Go()\n    Recalculate\nEnd Sub\n", second_exporter=_EXPORTS
    )
    assert _CODE in codes


def test_the_message_names_both_exporters() -> None:
    modules = [
        ModuleInput("Helpers", _STD, _EXPORTS),
        ModuleInput("Utils", _STD, _EXPORTS),
        ModuleInput("Caller", _STD, "Option Explicit\nSub Go()\n    Recalculate\nEnd Sub\n"),
    ]
    message = next(
        d.message for d in analyze_project(modules)["Caller"] if d.code == _CODE
    )
    assert "Helpers" in message and "Utils" in message


@pytest.mark.parametrize(
    ("label", "caller"),
    [
        ("qualified call", "Option Explicit\nSub Go()\n    Helpers.Recalculate\nEnd Sub\n"),
        (
            "caller declares it",
            "Option Explicit\nSub Recalculate()\nEnd Sub\nSub Go()\n    Recalculate\nEnd Sub\n",
        ),
        (
            "a local shadows it",
            "Option Explicit\nSub Go()\n    Dim Recalculate As Long\n    Recalculate = 1\nEnd Sub\n",
        ),
    ],
)
def test_silent_when_something_settles_the_name(label: str, caller: str) -> None:
    assert _CODE not in _caller_codes(caller, second_exporter=_EXPORTS)


def test_silent_when_only_one_module_exports_it() -> None:
    assert _CODE not in _caller_codes("Option Explicit\nSub Go()\n    Recalculate\nEnd Sub\n")


def test_a_private_procedure_is_not_exported_so_cannot_collide() -> None:
    assert _CODE not in _caller_codes(
        "Option Explicit\nSub Go()\n    Recalculate\nEnd Sub\n",
        second_exporter="Private Sub Recalculate()\nEnd Sub\n",
    )


def test_silent_without_project_context() -> None:
    """A single module in isolation cannot know what the project exports."""
    assert _CODE not in _codes("Option Explicit\nSub Go()\n    Recalculate\nEnd Sub\n")


# -- missing return assignment, widened to typed functions -----------------


def test_a_typed_function_that_never_assigns_its_return_reports() -> None:
    assert _codes("Public Function F() As Long\nEnd Function\n") == ["missing-return-assignment"]


@pytest.mark.parametrize(
    ("label", "body"),
    [
        ("bare assignment", "    F = 1"),
        ("assignment inside a single-line If", "    If x Then F = 1"),
        ("assignment in the Else of a single-line If", "    If x Then Debug.Print 1 Else F = 2"),
        ("a name that spells a keyword", None),
        ("a field of the returned UDT", None),
    ],
)
def test_silent_when_the_return_is_assigned(label: str, body: str | None) -> None:
    if body is not None:
        assert _codes(f"Public Function F() As Long\n{body}\nEnd Function\n") == []
        return
    if label == "a name that spells a keyword":
        # The lexer classifies `Read` as a keyword; the `=` is what settles it.
        assert _codes("Public Function Read() As Boolean\n    Read = True\nEnd Function\n") == []
    else:
        source = (
            "Private Type TPoint\n    X As Long\nEnd Type\n\n"
            "Private Function MakePoint() As TPoint\n    MakePoint.X = 1\nEnd Function\n"
        )
        assert _codes(source) == []


def test_a_body_that_raises_owes_no_return() -> None:
    source = (
        "Public Function F() As Long\n    Err.Raise 5, , \"nope\"\nEnd Function\n"
    )
    assert _codes(source) == []


def test_an_empty_member_of_an_interface_is_a_stub_not_unfinished_code() -> None:
    """A class another module declares with `Implements` states its members for
    the implementer to fill in, so every one of them is empty on purpose."""
    modules = [
        ModuleInput("ICallable", ModuleSymbolKind.CLASS, "Public Function Run() As Long\nEnd Function\n"),
        ModuleInput(
            "Worker",
            ModuleSymbolKind.CLASS,
            "Implements ICallable\nPrivate Function ICallable_Run() As Long\n    ICallable_Run = 1\nEnd Function\n",
        ),
    ]
    assert _project_codes(modules, "ICallable") == []


def test_an_empty_function_outside_an_interface_still_reports() -> None:
    modules = [
        ModuleInput("Lonely", ModuleSymbolKind.CLASS, "Public Function Run() As Long\nEnd Function\n"),
    ]
    assert _project_codes(modules, "Lonely") == ["missing-return-assignment"]


def test_a_property_get_is_covered_too() -> None:
    assert _codes("Public Property Get V() As Long\nEnd Property\n") == [
        "missing-return-assignment"
    ]


def test_a_sub_is_not_covered() -> None:
    assert _codes("Public Sub S()\nEnd Sub\n") == []
