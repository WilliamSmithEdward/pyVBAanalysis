"""Rules must not pair things from mutually exclusive `#If` arms.

Ports the XLIDE issue #58 cluster. Only one arm of an `#If` chain is ever built,
so two declarations in different arms are alternatives, not duplicates, and a
label in the other arm is not a `GoTo` target.

The arms only stay live when the compiler constant is UNKNOWN. With a known
constant the inactive arm is pruned before any rule sees it, which is why every
case here uses a constant the analyzer cannot evaluate.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_project
from pyvbaanalysis.conditional import create_conditional_activity_tracker
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions
from pyvbaanalysis.parser.nodes import Span
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind


def _codes(source: str, **kwargs: object) -> list[str]:
    return [
        d.code
        for d in analyze_module(source, AnalyzeModuleOptions(**kwargs))  # type: ignore[arg-type]
        if d.code != "option-explicit-missing"
    ]


# -- the primitive ---------------------------------------------------------

_ARMS = """#If CUSTOMFLAG Then
Dim A As Long
#Else
Dim B As Long
#End If

#If OTHERFLAG Then
Dim C As Long
#End If
Dim D As Long
"""


def _span_of(source: str, text: str) -> Span:
    start = source.index(text)
    return Span(start, start + len(text))


def test_mutually_exclusive_sees_only_arms_of_one_chain() -> None:
    tracker = create_conditional_activity_tracker(parse_module(_ARMS))
    assert tracker is not None
    a, b = _span_of(_ARMS, "Dim A"), _span_of(_ARMS, "Dim B")
    c, d = _span_of(_ARMS, "Dim C"), _span_of(_ARMS, "Dim D")
    # Two arms of one chain exclude each other.
    assert tracker.mutually_exclusive(a, b)
    # A span never excludes itself, and two SEPARATE chains do not exclude each
    # other: a build may take one and not the other.
    assert not tracker.mutually_exclusive(a, a)
    assert not tracker.mutually_exclusive(a, c)
    assert not tracker.mutually_exclusive(d, c)


def test_in_same_branch_is_stricter_than_not_exclusive() -> None:
    tracker = create_conditional_activity_tracker(parse_module(_ARMS))
    assert tracker is not None
    a = _span_of(_ARMS, "Dim A")
    c, d = _span_of(_ARMS, "Dim C"), _span_of(_ARMS, "Dim D")
    assert tracker.in_same_branch(a, a)
    # Neither exclusive nor the same branch: separate chains, and inside-vs-outside.
    assert not tracker.in_same_branch(a, c)
    assert not tracker.in_same_branch(d, c)


# -- the rules that pair declarations --------------------------------------


@pytest.mark.parametrize(
    ("label", "body"),
    [
        ("module member", "Public Const MODE As String = \"a\"\n#Else\nPublic Const MODE As String = \"b\""),
        ("procedure", "Public Sub Go()\nEnd Sub\n#Else\nPublic Sub Go()\nEnd Sub"),
        (
            "Declare",
            'Private Declare PtrSafe Sub S Lib "k" ()\n#Else\nPrivate Declare Sub S Lib "k" ()',
        ),
    ],
)
def test_one_name_per_arm_is_not_a_duplicate(label: str, body: str) -> None:
    assert _codes(f"Option Explicit\n\n#If CUSTOMFLAG Then\n{body}\n#End If\n") == []


def test_a_local_per_arm_is_not_a_duplicate_declaration() -> None:
    source = (
        "Option Explicit\n\nSub Go()\n#If CUSTOMFLAG Then\n    Dim v As Long\n"
        "#Else\n    Dim v As String\n#End If\nEnd Sub\n"
    )
    assert _codes(source) == []


def test_the_same_name_twice_in_one_arm_is_still_a_duplicate() -> None:
    """The carve-out must not swallow a genuine repeat inside a single arm."""
    source = (
        "Option Explicit\n\n#If CUSTOMFLAG Then\nPublic Const MODE As String = \"a\"\n"
        "Public Const MODE As String = \"b\"\n#End If\n"
    )
    assert _codes(source) == ["duplicate-module-variable"]


def test_duplicates_outside_any_directive_still_report() -> None:
    assert _codes("Option Explicit\nPublic Sub Go()\nEnd Sub\nPublic Sub Go()\nEnd Sub\n") == [
        "duplicate-procedure"
    ]


def test_distinct_property_accessors_still_share_their_name() -> None:
    source = (
        "Option Explicit\nPublic Property Get V() As Long\nV = 1\nEnd Property\n"
        "Public Property Let V(ByVal x As Long)\nEnd Property\n"
    )
    assert _codes(source) == []


# -- labels, For/Next, Else order ------------------------------------------


def test_a_label_in_another_arm_is_not_a_goto_target() -> None:
    source = (
        "Option Explicit\n\nSub Go()\n#If CUSTOMFLAG Then\n    GoTo Done\nDone:\n"
        "#Else\nDone:\n#End If\nEnd Sub\n"
    )
    assert "undefined-label" not in _codes(source)


def test_a_label_per_arm_is_not_a_duplicate_label() -> None:
    source = (
        "Option Explicit\n\nSub Go()\n#If CUSTOMFLAG Then\nDone:\n#Else\nDone:\n#End If\n"
        "End Sub\n"
    )
    assert "duplicate-label" not in _codes(source)


def test_a_genuinely_undefined_label_still_reports() -> None:
    assert "undefined-label" in _codes("Option Explicit\n\nSub Go()\n    GoTo Nowhere\nEnd Sub\n")


def test_for_and_next_in_different_chains_are_not_paired() -> None:
    """Each arm is internally consistent; the parser just sees one loop."""
    source = (
        "Option Explicit\n\nSub Go()\n    Dim i As Long\n    Dim j As Long\n"
        "#If CUSTOMFLAG Then\n    For i = 1 To 2\n#Else\n    For j = 1 To 2\n#End If\n"
        "    Next i\nEnd Sub\n"
    )
    assert "next-variable-mismatch" not in _codes(source)


def test_a_real_next_mismatch_still_reports() -> None:
    source = (
        "Option Explicit\n\nSub Go()\n    Dim i As Long\n    Dim j As Long\n"
        "    For i = 1 To 2\n    Next j\nEnd Sub\n"
    )
    assert "next-variable-mismatch" in _codes(source)


# -- Option and Implements placement ---------------------------------------


def test_a_const_directive_may_precede_option_explicit() -> None:
    """VBE-oracle verified (const_directive_before_option_explicit_compile): a
    conditional-compilation directive is not a declaration."""
    source = (
        "#Const XLIDEORACLEFLAG = 1\nOption Explicit\nPublic Sub Go()\n"
        "#If XLIDEORACLEFLAG Then\n    Debug.Print \"on\"\n#End If\nEnd Sub\n"
    )
    assert _codes(source) == []


def test_a_declaration_in_the_other_arm_closes_no_option_window() -> None:
    source = (
        "#If CUSTOMFLAG Then\nPublic X As Long\n#Else\nOption Explicit\n#End If\n"
    )
    assert "option-after-declaration" not in _codes(source)


def test_option_after_a_real_declaration_still_reports() -> None:
    assert "option-after-declaration" in _codes("Public X As Long\nOption Explicit\n")


def test_implements_after_a_procedure_in_another_arm_is_allowed() -> None:
    source = (
        "Option Explicit\n#If CUSTOMFLAG Then\nPublic Sub S()\nEnd Sub\n"
        "#Else\nImplements IFoo\n#End If\n"
    )
    assert "implements-statement-placement" not in _codes(
        source, module_kind=ModuleSymbolKind.CLASS
    )


def test_implements_after_a_real_procedure_still_reports() -> None:
    source = "Option Explicit\nPublic Sub S()\nEnd Sub\nImplements IFoo\n"
    assert "implements-statement-placement" in _codes(source, module_kind=ModuleSymbolKind.CLASS)


def test_a_project_pass_carries_the_same_behaviour() -> None:
    source = (
        "Option Explicit\n\n#If CUSTOMFLAG Then\nPublic Sub Go()\nEnd Sub\n"
        "#Else\nPublic Sub Go()\nEnd Sub\n#End If\n"
    )
    modules = [ModuleInput("Mod1", ModuleSymbolKind.STANDARD, source)]
    assert analyze_project(modules) == {"Mod1": []}
