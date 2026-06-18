"""M6: module-kind rule family (moduleKind.ts parity, self-contained slice)."""

from __future__ import annotations

import pytest
from oracle_support import accepted_cases, assert_oracle_behavior, asserted_cases, case_codes

from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.symbols import ModuleSymbolKind

_MK_CODES = (
    "object-module-public-member",
    "event-declaration-module-kind",
    "me-outside-object-module",
    "withevents-declaration",
    "friend-declaration",
    "implements-statement-placement",
    "raiseevent-undeclared-event",
    "declare-missing-ptrsafe",
)


def _codes(source: str, module_kind: ModuleSymbolKind = ModuleSymbolKind.STANDARD, **kw: object) -> set[str]:
    opts = AnalyzeModuleOptions(module_kind=module_kind, **kw)  # type: ignore[arg-type]
    return {d.code for d in analyze_module(source, opts)}


def test_object_module_public_members() -> None:
    assert "object-module-public-member" in _codes("Public Type T\n X As Long\nEnd Type", ModuleSymbolKind.CLASS)
    assert "object-module-public-member" in _codes("Public a() As Long", ModuleSymbolKind.CLASS)
    # Standard modules allow these public members.
    assert "object-module-public-member" not in _codes("Public Type T\n X As Long\nEnd Type", ModuleSymbolKind.STANDARD)


def test_event_declaration_module_kind() -> None:
    assert "event-declaration-module-kind" in _codes("Public Event Click()", ModuleSymbolKind.STANDARD)
    assert "event-declaration-module-kind" not in _codes("Public Event Click()", ModuleSymbolKind.CLASS)


def test_me_outside_object_module() -> None:
    assert "me-outside-object-module" in _codes("Sub S\n    Me.X = 1\nEnd Sub", ModuleSymbolKind.STANDARD)
    assert "me-outside-object-module" not in _codes("Sub S\n    Me.X = 1\nEnd Sub", ModuleSymbolKind.CLASS)
    # A member named Me (after a dot) is not the Me keyword.
    assert "me-outside-object-module" not in _codes("Sub S\n    obj.Me = 1\nEnd Sub", ModuleSymbolKind.STANDARD)


def test_with_events_declarations() -> None:
    assert "withevents-declaration" in _codes("Public WithEvents obj As Thing", ModuleSymbolKind.STANDARD)
    assert "withevents-declaration" not in _codes("Public WithEvents obj As Thing", ModuleSymbolKind.CLASS)
    # WithEvents As New / array are invalid even in a class.
    assert "withevents-declaration" in _codes("Public WithEvents obj As New Thing", ModuleSymbolKind.CLASS)


def test_friend_declarations() -> None:
    assert "friend-declaration" in _codes("Friend Sub S()\nEnd Sub", ModuleSymbolKind.STANDARD)
    assert "friend-declaration" not in _codes("Friend Sub S()\nEnd Sub", ModuleSymbolKind.CLASS)
    # Friend on a variable is always wrong.
    assert "friend-declaration" in _codes("Friend a As Long", ModuleSymbolKind.CLASS)


def test_implements_statement_placement() -> None:
    assert "implements-statement-placement" in _codes("Implements IFoo", ModuleSymbolKind.STANDARD)
    assert "implements-statement-placement" not in _codes("Implements IFoo", ModuleSymbolKind.CLASS)


def test_raise_event_targets() -> None:
    undeclared = "Sub S\n    RaiseEvent Click\nEnd Sub"
    assert "raiseevent-undeclared-event" in _codes(undeclared, ModuleSymbolKind.CLASS)
    declared = "Public Event Click()\nSub S\n    RaiseEvent Click\nEnd Sub"
    assert "raiseevent-undeclared-event" not in _codes(declared, ModuleSymbolKind.CLASS)


def test_declare_ptr_safe_for_win64() -> None:
    win64 = ConditionalCompilationEnvironment(compiler_constants={"win64": True})
    assert "declare-missing-ptrsafe" in _codes(
        'Declare Sub Beep Lib "k" ()', conditional_compilation=win64
    )
    assert "declare-missing-ptrsafe" not in _codes(
        'Declare PtrSafe Sub Beep Lib "k" ()', conditional_compilation=win64
    )
    # Without an explicit win64 the rule stays silent.
    assert "declare-missing-ptrsafe" not in _codes('Declare Sub Beep Lib "k" ()')


@pytest.mark.parametrize("code", _MK_CODES)
def test_oracle_asserted_cases(code: str) -> None:
    if not asserted_cases(code):
        pytest.skip(f"{code} has no asserted corpus cases")
    assert assert_oracle_behavior(code) > 0


def test_no_module_kind_false_positives_on_accepted_cases() -> None:
    mk = set(_MK_CODES)
    for case in accepted_cases():
        spurious = case_codes(case) & mk
        assert not spurious, f"{case.id}: module-kind false positive {spurious}"
