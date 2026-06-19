"""M7: object-variable-not-set (objectState.ts parity, dataflow-gated)."""

from __future__ import annotations

from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.diagnostics import analyze_module

_CODE = "object-variable-not-set"


def _codes(source: str) -> set[str]:
    return {d.code for d in analyze_module(source)}


def test_unset_object_member_access_fires() -> None:
    # The oracle case: a local Object is Nothing when a member is accessed.
    assert _CODE in _codes("Sub S\n    Dim obj As Object\n    obj.ToString\nEnd Sub")


def test_set_before_access_is_silent() -> None:
    assert _CODE not in _codes(
        "Sub S\n    Dim obj As Object\n    Set obj = Something\n    obj.ToString\nEnd Sub"
    )


def test_set_to_nothing_resets_to_unset() -> None:
    assert _CODE in _codes(
        "Sub S\n    Dim obj As Object\n"
        "    Set obj = Something\n    Set obj = Nothing\n    obj.ToString\nEnd Sub"
    )


def test_balanced_if_set_on_every_arm_is_silent() -> None:
    # Branch merge proves obj is set on every path before the access.
    src = (
        "Sub S\n    Dim obj As Object\n"
        "    If c Then\n        Set obj = A\n    Else\n        Set obj = B\n    End If\n"
        "    obj.ToString\nEnd Sub"
    )
    assert _CODE not in _codes(src)


def test_if_set_on_one_arm_only_is_silent_after_merge() -> None:
    # Set on only one arm -> merged state is 'unknown', not 'unset', so no false positive.
    src = (
        "Sub S\n    Dim obj As Object\n"
        "    If c Then\n        Set obj = A\n    End If\n"
        "    obj.ToString\nEnd Sub"
    )
    assert _CODE not in _codes(src)


def test_helper_call_argument_demotes_to_unknown() -> None:
    # Passing obj to a helper (possible ByRef rebind) makes its state unknown.
    src = "Sub S\n    Dim obj As Object\n    Init obj\n    obj.ToString\nEnd Sub"
    assert _CODE not in _codes(src)


def test_with_block_on_unset_object_fires() -> None:
    assert _CODE in _codes("Sub S\n    Dim obj As Object\n    With obj\n    End With\nEnd Sub")


def test_unstructured_flow_falls_back_to_conservative() -> None:
    # A GoTo makes the procedure unstructured: the branch merge is unsound, so the
    # straight-line walk demotes the set-on-both-arms obj to unknown (no fire, no FP).
    src = (
        "Sub S\n    Dim obj As Object\n"
        "    If c Then\n        Set obj = A\n    Else\n        Set obj = B\n    End If\n"
        "    GoTo Done\nDone:\n    obj.ToString\nEnd Sub"
    )
    assert _CODE not in _codes(src)


def test_non_object_locals_are_not_tracked() -> None:
    # A scalar local is not a tracked object variable.
    assert _CODE not in _codes("Sub S\n    Dim n As Long\n    n.ToString\nEnd Sub")


def test_oracle_asserted_cases() -> None:
    if asserted_cases(_CODE):
        assert assert_oracle_behavior(_CODE) > 0


def test_no_false_positives_on_accepted_cases() -> None:
    for case in accepted_cases():
        assert _CODE not in case_codes(case), f"{case.id}: {_CODE} false positive"
