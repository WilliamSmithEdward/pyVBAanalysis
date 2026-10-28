"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (rule fixes).

Each section names the upstream commit whose own tests it follows.
"""

import pytest

from pyvbaanalysis import analyze_module
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions


def _hits(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list:
    return [d for d in analyze_module(source, options) if d.code == code]


def _starts(source: str, code: str, options: AnalyzeModuleOptions | None = None) -> list[int]:
    return [d.span.start for d in _hits(source, code, options)]


# -- 53e6868d: the Else of a one-line If after a colon -----------------------

_ELSE_TAIL_QUIET = [
    "Sub T()\n    If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else: c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: _\n        Else c = 2\nEnd Sub\n",
    "Sub T()\n    If a Then For i = 1 To 2: Next: Else x = 1\n"
    "    If a Then With c: End With: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then Select Case b: Case 1: c = 1: End Select: Else d = 1\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: Else If x Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\nL1: If a Then b = 1: Else c = 2\n10 If a Then b = 1: Else c = 2\nEnd Sub\n",
    "Sub T()\n#If True Then\n    If a Then b = 1: Else c = 2\n#Else\n    Else\n#End If\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1 Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then If b Then c = 1: Else d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then b = 1: If c Then d = 2: Else e = 3\nEnd Sub\n",
    "Sub T()\n    If a Then obj.Else = 1: Else b\nEnd Sub\n",
]


@pytest.mark.parametrize("source", _ELSE_TAIL_QUIET)
def test_else_tail_after_colon_is_accepted(source: str) -> None:
    assert _hits(source, "else-without-if") == []


@pytest.mark.parametrize("source", [
    "Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2\n    Else\n        y = 1\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then If b Then c = 1 Else d = 2: Else e = 3\n    End If\nEnd Sub\n",
    "Sub T()\n    If x Then\n        If a Then b Else c: If d Then e: Else f\n    End If\nEnd Sub\n",
])
def test_else_tail_inside_block_if_is_not_a_clause(source: str) -> None:
    assert _hits(source, "else-without-if") == []
    assert _hits(source, "else-branch-order") == []


@pytest.mark.parametrize("source, marker", [
    ("Sub T()\n    If a Then b = 1: Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then b = 1 Else c = 2: Else d = 3\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: Else c = 2: Else d = 3\n"
     "    Else\n        y = 1\n    End If\nEnd Sub\n", "Else d"),
    ("Sub T()\n    Else\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1:\n    Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    b = 1: Else c = 2\nEnd Sub\n", "Else"),
    ("Sub T()\n    If a Then b = 1: ElseIf c Then d = 2\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If x Then\n        If a Then b = 1: ElseIf c Then d = 2\n    End If\nEnd Sub\n", "ElseIf"),
    ("Sub T()\n    If a Then b Else c: End If: Else d\nEnd Sub\n", "Else d"),
    ("Sub T()\n    If a Then obj.If = 1: Else b: Else c\nEnd Sub\n", "Else c"),
    ("Sub T()\n    If a Then b: Else c: Else d: If e Then f: Else g\nEnd Sub\n", "Else d"),
])
def test_stray_else_after_colon_is_still_reported(source: str, marker: str) -> None:
    assert _starts(source, "else-without-if") == [source.index(marker)]
    assert _hits(source, "else-branch-order") == []


def test_block_if_with_two_else_clauses_still_reported() -> None:
    source = "Sub T()\n    If x Then\n        a = 1\n    Else\n        b = 2\n    Else\n        c = 3\n    End If\nEnd Sub\n"
    assert len(_hits(source, "else-branch-order")) == 1
