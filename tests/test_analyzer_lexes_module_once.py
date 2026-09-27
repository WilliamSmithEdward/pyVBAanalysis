"""One analysis pass lexes the module once (XLIDE issue #139).

Mirrors upstream's tests/analyzerLexesModuleOnce.test.ts. The statement-token cache
is keyed by source string and holds two of them; the module token cache holds
eight. Rules that lex an expression the parser carried as its own string (an If
condition, a For Each source, an Enum member value, a parameter default, a Const
value) must lex it uncached. Sending those strings through the cached lexer evicted
the module, so the next ordinary statement re-lexed the whole module: once per
procedure, quadratic in module size, and 17x slower on real projects.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.lexer.tokenize import start_tokenize_miss_log_for_tests, stop_tokenize_miss_log_for_tests


def _build(procedures: int) -> str:
    lines = ["Option Explicit", "", "Public Enum Colours"]
    lines.extend(f"    Colour{i} = {i * 3}" for i in range(procedures))
    lines.extend(["End Enum", ""])
    for i in range(procedures):
        lines.extend(
            [
                f"Private Const LIMIT{i} As Long = {i} * 7 + 1",
                f"Public Sub Proc{i}(ByVal a As Long, Optional ByVal b As Integer = {i + 100})",
                "    Dim total As Long",
                "    Dim item As Variant",
                f"    total = a + LIMIT{i}",
                f"    If total <> {i} Then",
                "        total = b \\ total",
                f"    ElseIf a > {i + 1} Then",
                "        total = 0",
                "    End If",
                f"    For Each item In Array({i}, total)",
                "        total = total + item",
                "    Next item",
                "End Sub",
                "",
            ]
        )
    return "\r\n".join(lines)


def test_derived_expression_strings_do_not_evict_the_module() -> None:
    source = _build(60)
    start_tokenize_miss_log_for_tests()
    try:
        analyze_module(source, AnalyzeModuleOptions(host="excel"))
    finally:
        misses = stop_tokenize_miss_log_for_tests()
    assert [length for length in misses if length == len(source)] == [len(source)]
