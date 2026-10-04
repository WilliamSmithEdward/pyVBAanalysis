"""moduleState.ts parity: the names a module may write, and the module variables
nothing writes (XLIDE issue #241). Cases from upstream's moduleStateReuse,
moduleWriteCalleeScan and diagnostics/moduleState241 tests."""

from __future__ import annotations

from pyvbaanalysis.diagnostics.module_state import (
    remember_project_written_names,
    untouched_module_variables,
    untouched_module_variables_in,
    written_names_in,
)
from pyvbaanalysis.parser.nodes import ProcedureNode
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind, ModuleSymbols


def _fixture(source: str) -> tuple[ModuleSymbols, list[ProcedureNode]]:
    mod = parse_module(source)
    symbols = build_module_symbols(
        "M", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=mod)
    )
    return symbols, [member for member in mod.members if isinstance(member, ProcedureNode)]


def test_written_names_take_targets_writing_heads_and_whole_names_passed() -> None:
    names = written_names_in(
        "\n".join(
            [
                "Sub T()",
                "    a = 1",
                "    Set b = Nothing",
                "    ReDim c(3)",
                "    Fill d",
                "    Call Fill(e)",
                "    x = Take(f) + CLng(g)",
                "    If n = 1 Then h = 2",
                "    Debug.Print i",
                "    For j = 1 To 2: Next",
                "End Sub",
            ]
        )
    )
    for name in ["a", "b", "c", "d", "e", "f", "h", "j", "x"]:
        assert name in names, name
    # Read only: a library function's argument, a condition, a print.
    for name in ["g", "n", "i"]:
        assert name not in names, name
    # The module's own Trim is no library function, and a parameter is no write.
    shadowed = written_names_in(
        "Sub T()\n    y = Trim(k)\nEnd Sub\nFunction Trim(x As String) As String\nEnd Function\nSub U(p)\nEnd Sub\n"
    )
    assert "k" in shadowed
    assert "p" not in shadowed


def test_written_names_keep_nested_runtime_and_unknown_callees_apart() -> None:
    assert list(written_names_in("Fill(Abs(x), Inner(y), z)")) == ["y", "z"]
    assert list(written_names_in("result = Abs(x) + Inner(y)")) == ["result", "y"]


def test_written_names_of_bare_calls_receivers_and_parenthesized_arguments() -> None:
    assert list(written_names_in("Fill x, y\nobj.Fill z, q")) == ["x", "y", "z", "q"]
    assert list(written_names_in("Fill (x), y")) == ["x", "y"]


def test_written_names_see_source_procedures_shadow_runtime_functions() -> None:
    assert list(written_names_in("Sub Abs()\nEnd Sub\nCall Abs(x, y)")) == ["x", "y"]
    assert list(written_names_in("Call Abs(x, y)")) == []


def test_written_names_stay_conservative_on_unmatched_parentheses() -> None:
    assert list(written_names_in("Fill(x, y")) == ["x", "y"]
    assert list(written_names_in("Fill(x)), y")) == ["x", "y"]
    assert list(written_names_in("Fill(Abs(x), Inner(y")) == ["y"]


def test_written_names_read_a_wide_call() -> None:
    count = 1000
    names = written_names_in("Fill(" + ",".join(f"v{i}" for i in range(count)) + ")")
    assert len(names) == count
    assert "v0" in names
    assert f"v{count - 1}" in names


def test_untouched_variables_keep_writes_exclusions_and_local_shadowing() -> None:
    source = "\n".join(
        [
            "Private a As Long",
            "Private b As Long",
            "Private changed As Long",
            "Private auto As New Collection",
            "Private fixed As String * 3",
            "Public shared As Long",
            "Sub LocalScope()",
            "Dim A As Long",
            "changed = 1",
            "End Sub",
            "Sub ParamScope(ByVal B As Long)",
            "x = 1",
            "End Sub",
            "Sub Unshadowed()",
            "x = 1",
            "End Sub",
        ]
    )
    symbols, procs = _fixture(source)
    remember_project_written_names(symbols, set())
    for _consumer in range(3):
        assert list(untouched_module_variables_in(source, symbols, procs[0])) == ["b", "shared"]
        assert list(untouched_module_variables_in(source, symbols, procs[1])) == ["a", "shared"]
        assert list(untouched_module_variables_in(source, symbols, procs[2])) == ["a", "b", "shared"]


def test_public_state_follows_the_project_writes() -> None:
    source = "Public shared As Long\nSub P()\nx = 1\nEnd Sub"
    symbols, procs = _fixture(source)

    def names() -> list[str]:
        return list(untouched_module_variables_in(source, symbols, procs[0]))

    assert names() == []
    writes: set[str] = set()
    remember_project_written_names(symbols, writes)
    assert names() == ["shared"]
    writes.add("shared")
    remember_project_written_names(symbols, writes)
    assert names() == []
    remember_project_written_names(symbols, set())
    assert names() == ["shared"]
    remember_project_written_names(symbols, None)
    assert names() == []


def test_module_write_facts_follow_edited_source() -> None:
    source = "Private a As Long\nSub P()\nx = 1\nEnd Sub"
    symbols, procs = _fixture(source)
    assert list(untouched_module_variables_in(source, symbols, procs[0])) == ["a"]
    edited = source.replace("x = 1", "a = 1")
    assert list(untouched_module_variables_in(edited, symbols, procs[0])) == []
    assert list(untouched_module_variables(source, symbols)) == ["a"]
