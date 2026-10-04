"""functionResults.ts and calleeArguments.ts parity: what a Function of the module
returns where its text fixes it (XLIDE issues #448, #562), and what a callee does
with an argument passed whole (issues #449, #665, #685)."""

from __future__ import annotations

from pyvbaanalysis.diagnostics.callee_arguments import callee_keeps_argument, callee_member_calls
from pyvbaanalysis.diagnostics.function_results import (
    FunctionNullResult,
    FunctionNumberResult,
    FunctionStringResult,
    function_integer_result,
    function_result_at,
    function_result_named,
    known_function_results,
)
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens
from pyvbaanalysis.parser.nodes import ModuleNode, ProcedureNode
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind, ModuleSymbols

_SOURCE = "\n".join(
    [
        "Function Zero() As Long",
        "End Function",
        "Function S() As String",
        '    S = "abc"',
        "End Function",
        "Function V()",
        "End Function",
        "Function H() As Integer",
        "    H = 2.5",
        "End Function",
        "Function R() As Long",
        "    If x Then Exit Function",
        "    R = 5",
        "End Function",
        "Function B() As Boolean",
        "    B = 5",
        "End Function",
        "Function Big() As Byte",
        "    Big = 300",
        "End Function",
        "Function Nul()",
        "    Nul = Null",
        "End Function",
        "Function Twice() As Long",
        "    Twice = 1",
        "    Twice = 2",
        "End Function",
        "Function Raises() As Long",
        "    Err.Raise 5",
        "    Raises = 1",
        "End Function",
        "Sub Caller()",
        "    x = Zero() + 1",
        "End Sub",
        "Sub Shadows(Zero As Long)",
        "End Sub",
    ]
)


def _setup() -> tuple[ModuleNode, ModuleSymbols, dict[str, ProcedureNode]]:
    mod = parse_module(_SOURCE)
    symbols = build_module_symbols(
        "M", ModuleSymbolKind.STANDARD, _SOURCE, BuildModuleSymbolsOptions(parsed_module=mod)
    )
    procs = {member.name.lower(): member for member in mod.members if isinstance(member, ProcedureNode)}
    return mod, symbols, procs


def test_known_function_results() -> None:
    mod, _symbols, _procs = _setup()
    results = known_function_results(_SOURCE, mod, None)
    assert results["zero"] == FunctionNumberResult(value=0, type="long")
    assert results["s"] == FunctionStringResult(value="abc")
    assert "v" not in results
    assert results["h"] == FunctionNumberResult(value=2, type="integer")
    assert "r" not in results
    assert results["b"] == FunctionNumberResult(value=-1, type="boolean")
    assert "big" not in results
    assert results["nul"] == FunctionNullResult()
    assert "twice" not in results
    assert "raises" not in results
    assert known_function_results(_SOURCE, mod, None) is results


def test_function_result_lookups_respect_shadowing() -> None:
    mod, symbols, procs = _setup()
    results = known_function_results(_SOURCE, mod, None)
    toks = raw_expression_tokens("Zero() + 1")
    found = function_result_at(toks, 0, results, procs["caller"], symbols)
    assert found is not None and found.end == 2 and found.result == results["zero"]
    assert function_result_at(raw_expression_tokens("Zero(1)"), 0, results, procs["caller"], symbols) is None
    assert function_result_at(toks, 0, results, procs["shadows"], symbols) is None
    assert function_result_named("s", results, procs["caller"], symbols) == results["s"]
    assert function_result_named("zero", results, procs["shadows"], symbols) is None
    assert function_integer_result("Zero()", results, procs["caller"], symbols) == 0
    assert function_integer_result("zero", results, procs["caller"], symbols) == 0
    assert function_integer_result("S", results, procs["caller"], symbols) is None


_CALLEES = "\n".join(
    [
        "Sub Reads(p As Collection)",
        "    Debug.Print TypeName(p)",
        "End Sub",
        "Sub Writes(p As Collection)",
        "    Set p = New Collection",
        "End Sub",
        "Sub ByValue(ByVal p As Collection)",
        "    Set p = Nothing",
        "End Sub",
        "Sub PassesOn(p As Collection)",
        "    x = Other(1, p)",
        "End Sub",
        "Sub Loops(p As Long)",
        "    For p = 1 To 2",
        "    Next",
        "End Sub",
        "Sub R1(ByVal p As Collection)",
        "    Dim i As Long",
        "    p.Remove 1",
        '    p.Add "a", Key:="k"',
        "End Sub",
        "Sub R2(ByVal p As Collection)",
        "    p.Add x",
        "End Sub",
    ]
)


def test_callee_keeps_argument() -> None:
    keeps = callee_keeps_argument(_CALLEES)
    assert keeps("Reads", 0) is True
    assert keeps("Writes", 0) is False
    assert keeps("ByValue", 0) is True
    assert keeps("PassesOn", 0) is False
    assert keeps("Loops", 0) is False
    assert keeps("Reads", 0, "p") is True
    assert keeps("Reads", 1) is False
    assert keeps("Missing", 0) is False


def test_callee_member_calls_replay_on_the_callers_name() -> None:
    member_calls = callee_member_calls(_CALLEES)
    replayed = member_calls(raw_expression_tokens("R1 c"))
    assert [" ".join(tok.raw_text for tok in stmt) for stmt in replayed["c"]] == [
        "c . Remove 1",
        'c . Add "a" , Key := "k"',
    ]
    assert dict(member_calls(raw_expression_tokens("Call R1(c)"))).keys() == {"c"}
    assert member_calls(raw_expression_tokens("R1 (c)")) == {}
    assert member_calls(raw_expression_tokens("R2 c")) == {}
