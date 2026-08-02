"""late-bound-friend-member (lateBinding.ts parity, xlide issue #2 / port issue #5).

Friend members are not on a class's IDispatch interface, so reaching one
through a Variant/Object receiver raises runtime error 438 while compiling
clean. Fires only on member names that resolve EXCLUSIVELY to Friend members
of exhaustive project class modules.
"""

from __future__ import annotations

from oracle_support import (
    accepted_cases,
    assert_oracle_behavior,
    asserted_cases,
    oracle_false_positives,
)

from pyvbaanalysis import analyze_project
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

_CODE = "late-bound-friend-member"

# A class whose ToastSlot is Friend - the shape from the real-world report.
_UI_CLASS = (
    "Option Explicit\n"
    "Private mSlot As Long\n"
    "Friend Property Get ToastSlot() As Long\n"
    "    ToastSlot = mSlot\n"
    "End Property\n"
    "Friend Property Let ToastSlot(ByVal value As Long)\n"
    "    mSlot = value\n"
    "End Property\n"
    "Public Sub Show()\n"
    "End Sub\n"
)


def _analyze(source: str, extra: list[tuple[str, ModuleSymbolKind, str]] | None = None):
    modules = [
        ModuleInput("Caller", ModuleSymbolKind.STANDARD, source),
        ModuleInput("ReDimUI", ModuleSymbolKind.CLASS, _UI_CLASS),
        *(ModuleInput(name, kind, src) for name, kind, src in (extra or [])),
    ]
    return [d for d in analyze_project(modules)["Caller"] if d.code == _CODE]


def test_flags_friend_member_through_collection_element() -> None:
    # The reported bug: Collection.Item returns Variant, so the element is
    # late bound and cannot see ToastSlot. Compiles clean; dies at run time.
    src = (
        "Option Explicit\n"
        "Public Sub Compact(ByVal live As Collection)\n"
        "    Dim position As Long\n"
        "    For position = 2 To live.Count\n"
        "        Debug.Print live.Item(position).ToastSlot\n"
        "    Next position\n"
        "End Sub\n"
    )
    hits = _analyze(src)
    assert len(hits) == 1
    assert src[hits[0].span.start : hits[0].span.end] == "ToastSlot"
    assert "438" in hits[0].message
    assert "Friend member of class 'ReDimUI'" in hits[0].message


def test_flags_default_member_subscript_form() -> None:
    src = (
        "Option Explicit\n"
        "Public Sub Compact(ByVal live As Collection)\n"
        "    Debug.Print live(1).ToastSlot\n"
        "End Sub\n"
    )
    hits = _analyze(src)
    assert len(hits) == 1
    assert src[hits[0].span.start : hits[0].span.end] == "ToastSlot"


def test_flags_variant_object_and_untyped_locals() -> None:
    for declaration in ("Dim item As Variant", "Dim item As Object", "Dim item"):
        src = (
            "Option Explicit\n"
            "Public Sub T()\n"
            f"    {declaration}\n"
            "    Debug.Print item.ToastSlot\n"
            "End Sub\n"
        )
        hits = _analyze(src)
        assert len(hits) == 1, declaration
        assert src[hits[0].span.start : hits[0].span.end] == "ToastSlot", declaration


def test_silent_once_assigned_to_typed_local() -> None:
    # The documented fix: the typed intermediate makes the call early bound.
    src = (
        "Option Explicit\n"
        "Public Sub Compact(ByVal live As Collection)\n"
        "    Dim candidate As ReDimUI\n"
        "    Set candidate = live.Item(1)\n"
        "    Debug.Print candidate.ToastSlot\n"
        "End Sub\n"
    )
    assert _analyze(src) == []


def test_silent_when_name_is_also_public_somewhere() -> None:
    # Another class exposes the same name publicly, so a late-bound receiver
    # could legally dispatch there. The runtime type is unknowable; say nothing.
    other = "Option Explicit\nPublic Property Get ToastSlot() As Long\nEnd Property\n"
    src = (
        "Option Explicit\n"
        "Public Sub T()\n"
        "    Dim item As Variant\n"
        "    Debug.Print item.ToastSlot\n"
        "End Sub\n"
    )
    assert _analyze(src, [("OtherThing", ModuleSymbolKind.CLASS, other)]) == []


def test_silent_for_host_member_name() -> None:
    # A Friend member named Value collides with the host surface, so a
    # late-bound receiver may well be a Range.
    shadowing = "Option Explicit\nFriend Property Get Value() As Long\nEnd Property\n"
    src = (
        "Option Explicit\n"
        "Public Sub T()\n"
        "    Dim item As Variant\n"
        "    Debug.Print item.Value\n"
        "End Sub\n"
    )
    modules = [
        ModuleInput("Caller", ModuleSymbolKind.STANDARD, src),
        ModuleInput("Shadowing", ModuleSymbolKind.CLASS, shadowing),
    ]
    assert [d for d in analyze_project(modules)["Caller"] if d.code == _CODE] == []


def test_silent_for_unknown_member_on_late_bound_receiver() -> None:
    # The VBE oracle records this as compile-valid: the runtime type could
    # support anything, so an unknown name is never evidence of a defect.
    src = (
        "Option Explicit\n"
        "Public Sub T()\n"
        "    Dim item As Variant\n"
        "    Debug.Print item.NoSuchMemberAnywhere\n"
        "End Sub\n"
    )
    assert _analyze(src) == []


def test_silent_on_strongly_typed_receiver() -> None:
    # Friend is legal within the project when the call is early bound.
    src = (
        "Option Explicit\n"
        "Public Sub T()\n"
        "    Dim ui As ReDimUI\n"
        "    Set ui = New ReDimUI\n"
        "    Debug.Print ui.ToastSlot\n"
        "End Sub\n"
    )
    assert _analyze(src) == []


def test_silent_with_no_project_class_context() -> None:
    src = (
        "Option Explicit\n"
        "Public Sub T()\n"
        "    Dim item As Variant\n"
        "    Debug.Print item.ToastSlot\n"
        "End Sub\n"
    )
    modules = [ModuleInput("Caller", ModuleSymbolKind.STANDARD, src)]
    assert [d for d in analyze_project(modules)["Caller"] if d.code == _CODE] == []


def test_fires_inside_the_owning_class_itself() -> None:
    # Being inside the class does not put Friend members back on the dispatch
    # interface, so an Object-typed local fails here too (VBE oracle
    # late_bound_friend_member_same_class_runtime).
    probe = (
        "Public Function Probe(ByVal value As Variant) As Long\n"
        "    Dim candidate As Object\n"
        "    Set candidate = value\n"
        "    If TypeOf candidate Is ReDimUI Then\n"
        "        Probe = candidate.ToastSlot\n"
        "    End If\n"
        "End Function\n"
    )
    self_class = _UI_CLASS + probe
    modules = [ModuleInput("ReDimUI", ModuleSymbolKind.CLASS, self_class)]
    hits = [d for d in analyze_project(modules)["ReDimUI"] if d.code == _CODE]
    assert len(hits) == 1
    assert self_class[hits[0].span.start : hits[0].span.end] == "ToastSlot"


def test_silent_on_collection_member_itself() -> None:
    src = (
        "Option Explicit\n"
        "Public Sub T(ByVal live As Collection)\n"
        "    Debug.Print live.Count\n"
        "End Sub\n"
    )
    assert _analyze(src) == []


def test_oracle_asserted_cases() -> None:
    assert asserted_cases(_CODE)
    assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        spurious = oracle_false_positives(case, (_CODE,))
        assert not spurious, f"{case.id}: late-binding false positive {spurious}"
