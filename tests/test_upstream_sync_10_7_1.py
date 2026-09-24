"""Behaviour added by the sync to XLIDE 10.7.1 (issues #82 to #90).

Each test mirrors one of upstream's own, and upstream's recorded calls for them
replay through the port unchanged.

* What a one-line If runs after a colon belongs to the If (MS-VBAL 5.4.2.9), and
  `If x Then:` is a one-line If (#84).
* A comment ending in ` _` runs on to the next line (#82), and a continuation may
  have whitespace after its `_` (#83).
* A comma right after a callee opens the argument list (#85).
* Contextual keywords are keywords only inside their own statement (#86).
* `1.`, `&17`, `=>`, `=<`, `><` and `a < > b` read as the VBE reads them (#87).
* The one-word `EndIf` closes a block If (#88).

The corrected Excel return types (#90) are in test_rules_assignments_types.py.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis.conditional.conditional_compilation import evaluate_conditional_expression
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.lexer import TokenKind, VbaToken, tokenize
from pyvbaanalysis.lexer.token_kinds import Trivia, TriviaKind
from pyvbaanalysis.parser import parse_module
from pyvbaanalysis.parser.nodes import IfBlockNode, ProcedureNode


def _raws(tokens: list[VbaToken]) -> list[str]:
    return [t.raw_text for t in tokens]


def _kinds(tokens: list[VbaToken]) -> list[str]:
    return [t.kind.value for t in tokens]


def _round_trip(tokens: list[VbaToken]) -> str:
    parts: list[str] = []
    for token in tokens:
        parts.extend(trivia.text for trivia in token.leading_trivia)
        parts.append(token.raw_text)
    if tokens:
        parts.extend(trivia.text for trivia in tokens[-1].trailing_trivia)
    return "".join(parts)


def _codes(source: str, code: str) -> list[str]:
    return [source[d.span.start : d.span.end] for d in analyze_module(source) if d.code == code]


def _wrap(*lines: str) -> str:
    body = "".join(f"    {line}\n" for line in lines)
    return f"Public Sub T(ByVal x As Boolean)\n    Dim y As Long\n{body}End Sub\n"


# -- the lexer ------------------------------------------------------------------


def test_a_contextual_keyword_is_a_keyword_only_inside_its_statement() -> None:
    def words(source: str) -> list[str]:
        return [
            f"{t.raw_text}:{'k' if t.kind is TokenKind.KEYWORD else 'i'}"
            for t in tokenize(source)
            if t.kind in (TokenKind.KEYWORD, TokenKind.IDENTIFIER)
        ]

    assert words("Option Compare Text") == ["Option:k", "Compare:k", "Text:k"]
    assert words("Option Explicit: Option Base 1") == ["Option:k", "Explicit:k", "Option:k", "Base:k"]
    assert words('Declare PtrSafe Function F Lib "k" Alias "G" (ByVal text As Long)') == [
        "Declare:k", "PtrSafe:k", "Function:k", "F:i", "Lib:k", "Alias:k",
        "ByVal:k", "text:i", "As:k", "Long:k",
    ]
    assert words("For i = step To 10 Step step") == ["For:k", "i:i", "step:i", "To:k", "Step:k", "step:i"]
    assert words("On Error Resume Next: Error 5: x = Error(5)") == [
        "On:k", "Error:k", "Resume:k", "Next:k", "Error:k", "x:i", "Error:i",
    ]
    assert words("Open output For Output Access Read As #1") == [
        "Open:k", "output:i", "For:k", "Output:k", "Access:i", "Read:k", "As:k",
    ]
    assert words("x = r.Text & text & binary & read & append & random & base & compare & explicit & lib & alias") == [
        "x:i", "r:i", "Text:i", "text:i", "binary:i", "read:i", "append:i", "random:i",
        "base:i", "compare:i", "explicit:i", "lib:i", "alias:i",
    ]
    text = tokenize("Dim text As String")[1]
    assert (text.kind, text.canonical_text) == (TokenKind.IDENTIFIER, None)


def test_a_comment_ending_in_an_underscore_runs_on_through_the_next_line() -> None:
    # The VBE returns 1 from a function whose `n = 2` follows `' comment _`.
    source = "n = 1\r\n' comment _\r\nn = 2\r\nR = n"
    tokens = tokenize(source)
    assert _raws([t for t in tokens if t.kind is not TokenKind.NEWLINE]) == [
        "n", "=", "1", "' comment _\r\nn = 2", "R", "=", "n",
    ]
    after = next(t for t in tokens if t.raw_text == "R")
    assert (after.line, after.character) == (3, 0)
    assert _round_trip(tokens) == source


def test_a_trailing_comment_a_rem_comment_and_trailing_whitespace_all_continue() -> None:
    assert tokenize("n = 1 ' note _\n    this is not code\nn = 2")[3].raw_text == "' note _\n    this is not code"
    assert tokenize("Rem note _\nEnd Sub\nx = 1")[0].raw_text == "Rem note _\nEnd Sub"
    assert tokenize("' a _  \nb _\nc\nd")[0].raw_text == "' a _  \nb _\nc"
    following = next(t for t in tokenize("x = 1 ' a _\n  b\n  y = 2") if t.raw_text == "y")
    assert (following.line, following.character) == (2, 2)


def test_a_comment_ends_at_the_line_when_its_underscore_is_no_continuation() -> None:
    assert tokenize("' a_\nb")[0].raw_text == "' a_"
    assert tokenize("'_\nb")[0].raw_text == "'_"
    assert tokenize("' a _b\nc")[0].raw_text == "' a _b"


@pytest.mark.parametrize("source", ["1.", "1.#", "1.e5"])
def test_a_float_may_leave_out_its_fractional_digits(source: str) -> None:
    # The VBE stores `1.` as `1#`, `1.#` as `1#` and `1.e5` as `100000#`.
    assert [(t.kind, t.raw_text) for t in tokenize(source)] == [(TokenKind.FLOAT_LITERAL, source)]


def test_a_letter_after_the_dot_leaves_it_a_member_access() -> None:
    assert _raws(tokenize("x = Array(1., 2)")) == ["x", "=", "Array", "(", "1.", ",", "2", ")"]
    assert _kinds(tokenize("1.Value")) == ["integerLiteral", "punctuation", "identifier"]


def test_an_octal_literal_may_leave_out_its_o_wherever_it_stands() -> None:
    assert [(t.kind, t.raw_text) for t in tokenize("&17")] == [(TokenKind.INTEGER_LITERAL, "&17")]
    assert _raws(tokenize("x = &17&")) == ["x", "=", "&17&"]
    assert _raws(tokenize('x = "a" &1')) == ["x", "=", '"a"', "&1"]
    assert _raws(tokenize("x = &18")) == ["x", "=", "&1", "8"]
    # 8 and 9 are no octal digits, so `"a" &9` is a concatenation.
    assert _kinds(tokenize('"a" &9')) == ["stringLiteral", "operator", "integerLiteral"]


@pytest.mark.parametrize(("written", "standard"), [("=>", ">="), ("=<", "<="), ("><", "<>")])
def test_a_relational_operator_written_the_other_way_round_is_one(written: str, standard: str) -> None:
    operator = tokenize(f"a {written} b")[1]
    assert (operator.kind, operator.raw_text, operator.canonical_text) == (TokenKind.OPERATOR, written, standard)
    assert tokenize("a >= b")[1].canonical_text is None


def test_a_continuation_may_have_whitespace_after_its_underscore() -> None:
    # The VBE returns 3 from `R = 1 + _   ` / `2`, and drops the spaces.
    source = "R = 1 + _ \t \r\n    2"
    tokens = tokenize(source)
    assert _kinds(tokens) == ["identifier", "operator", "integerLiteral", "operator", "integerLiteral"]
    assert list(tokens[4].leading_trivia) == [
        Trivia(kind=TriviaKind.LINE_CONTINUATION, text=" _ \t \r\n", start=7, end=14),
        Trivia(kind=TriviaKind.WHITESPACE, text="    ", start=14, end=18),
    ]
    assert (tokens[4].line, tokens[4].character) == (1, 4)
    assert _round_trip(tokens) == source


def test_an_underscore_followed_by_text_or_the_end_is_a_token() -> None:
    assert _raws(tokenize("a _ b")) == ["a", "_", "b"]
    assert _raws(tokenize("a _  ")) == ["a", "_"]


# -- the parser -------------------------------------------------------------------


def test_if_then_colon_is_a_one_line_if() -> None:
    module = parse_module("Sub F()\n    If x Then:\n    y = 1\nEnd Sub\n")
    assert module.diagnostics == []
    procedure = module.members[0]
    assert isinstance(procedure, ProcedureNode)
    assert not any(isinstance(node, IfBlockNode) for node in procedure.body)
    # The VBE refuses an End If after it: "End If without block If".
    closed = parse_module("Sub F()\n    If x Then:\n    y = 1\n    End If\nEnd Sub\n")
    assert [d.message for d in closed.diagnostics] == ["Unexpected 'End If' without a matching opening block."]


def test_the_statements_a_one_line_if_runs_after_a_colon_are_marked() -> None:
    source = (
        "Sub F()\n"
        "    If x Then y = 1: z = 2: Exit Sub\n"
        "    If x Then:\n"
        "    w = 3\n"
        "    a = 1: If x Then b = 2 Else c = 3: d = 4\n"
        "End Sub\n"
    )
    procedure = parse_module(source).members[0]
    assert isinstance(procedure, ProcedureNode)
    assert [
        (source[node.span.start : node.span.end], bool(getattr(node, "single_line_if_tail", False)))
        for node in procedure.body
    ] == [
        ("If x Then y = 1", False),
        ("z = 2", True),
        ("Exit Sub", True),
        ("If x Then", False),
        ("w = 3", False),
        ("a = 1", False),
        ("If x Then b = 2 Else c = 3", False),
        ("d = 4", True),
    ]


def test_the_one_word_endif_closes_a_block_if() -> None:
    module = parse_module("Sub F()\n    If x Then\n        Exit Sub\n    EndIf\n    y = 2\nEnd Sub\n")
    assert module.diagnostics == []
    procedure = module.members[0]
    assert isinstance(procedure, ProcedureNode)
    block = next(node for node in procedure.body if isinstance(node, IfBlockNode))
    assert block.closed
    assert [type(node).__name__ for node in procedure.body] == ["IfBlockNode", "AssignmentNode"]


# -- the analysis -------------------------------------------------------------------


@pytest.mark.parametrize(
    "source",
    [
        _wrap("If x Then y = 1: Exit Sub", "Debug.Print 1"),
        _wrap("If x Then: Exit Sub", "Debug.Print 1"),
        _wrap("If x Then y = 1 Else y = 2: Exit Sub", "Debug.Print 1"),
        _wrap("If x Then y = 1: Exit Sub Else y = 3", "Debug.Print 1"),
        _wrap("If x Then y = 1: GoTo Done", "Debug.Print 1", "Done:"),
        # As a user reported it.
        "Public Function ACCT(ByVal conn As Object) As Variant\n"
        "    If conn Is Nothing Then ACCT = CVErr(xlErrNA): Exit Function\n"
        '    ACCT = conn.Execute("select 1").Fields(0).Value\n'
        "End Function\n",
    ],
)
def test_what_a_one_line_if_runs_after_a_colon_is_conditional(source: str) -> None:
    assert _codes(source, "unreachable-code") == []


def test_code_after_an_exit_under_if_then_colon_is_still_unreachable() -> None:
    source = _wrap("If x Then:", "Exit Sub", "Debug.Print 1")
    assert _codes(source, "unreachable-code") == ["Debug.Print 1"]


def test_code_after_a_block_closed_by_endif_is_reachable() -> None:
    assert _codes(_wrap("If x Then", "    Exit Sub", "EndIf", "y = 2"), "unreachable-code") == []


@pytest.mark.parametrize(
    "source",
    [
        "Option Explicit\nSub Test()\n    ' disabled: _\n    End Sub\n    Debug.Print 1\nEnd Sub\n",
        "Option Explicit\nSub Test()\n    Rem disabled: _\n    End Sub\n    Debug.Print 1\nEnd Sub\n",
        "Option Explicit\nSub Test()\n    Debug.Print 0 ' disabled: _\n    End Sub\n    Debug.Print 1\nEnd Sub\n",
        "Option Explicit\nSub Test()\n#If True Then ' note _\n    End Sub\n#End If\n    Debug.Print 1\nEnd Sub\n",
    ],
)
def test_the_line_after_a_continued_comment_is_comment_text(source: str) -> None:
    assert [d.code for d in analyze_module(source)] == []


def test_the_literal_and_operator_spellings_the_vbe_reads_are_accepted() -> None:
    # Measured in Excel 16.0: each line compiles.
    source = (
        "Option Explicit\n"
        "Sub A()\n"
        "    Dim a As Double, b As Double\n"
        "    a = 1.\n"
        "    a = &17\n"
        "    a = Array(1., &17&)(0)\n"
        "    If a => b Then a = 1\n"
        "    If a =< b Then a = 1\n"
        "    If a >< b Then a = 1\n"
        "    If a = > b Then a = 1\n"
        "    If a < > b Then a = 1\n"
        "    If a > = b Then a = 1\n"
        "End Sub\n"
    )
    assert _codes(source, "invalid-expression-syntax") == []


def test_an_octal_literal_after_a_complete_expression_is_refused() -> None:
    assert _codes('Sub T()\n    s = "a" &1\nEnd Sub\n', "invalid-expression-syntax") == ["&1"]
    assert _codes('Sub T()\n    s = "a" &9\nEnd Sub\n', "invalid-expression-syntax") == []


def test_an_underscore_with_whitespace_after_it_continues_the_line() -> None:
    source = (
        "Option Explicit\n"
        "Sub T()\n"
        "    Dim a As Boolean, _ \n"
        "        b As Boolean\n"
        "    If a And _  \n"
        "       b Then\n"
        '        Debug.Print "both"\n'
        "    End If\n"
        "End Sub\n"
    )
    assert [d.code for d in analyze_module(source)] == []


@pytest.mark.parametrize(
    "source",
    [
        "Public Sub T()\n    Dim obj As Object\n    If Ready Then Set obj = New Collection\n    obj.ToString\nEnd Sub\n",
        "Public Sub T()\n    Dim obj As Object\n    If Ready Then n = 1: Set obj = New Collection\n    obj.ToString\nEnd Sub\n",
        "Public Sub T()\n    Dim obj As Object\n    Set obj = New Collection\n"
        "    If Ready Then n = 1: Set obj = Nothing\n    obj.ToString\nEnd Sub\n",
    ],
)
def test_a_set_in_a_one_line_if_is_conditional(source: str) -> None:
    assert _codes(source, "object-variable-not-set") == []


def test_what_a_one_line_if_runs_after_its_colon_is_checked_on_its_path() -> None:
    never_set = "Public Sub T()\n    Dim obj As Object\n    If Ready Then n = 1: obj.ToString\nEnd Sub\n"
    assert _codes(never_set, "object-variable-not-set") == ["obj"]
    for source in (
        "Public Sub T()\n    Dim obj As Object\n    If Ready Then n = 1: Set obj = New Collection: obj.ToString\nEnd Sub\n",
        "Public Sub T()\n    Dim obj As Object\n    If Ready Then Set obj = New Collection: obj.ToString\nEnd Sub\n",
    ):
        assert _codes(source, "object-variable-not-set") == [], source


@pytest.mark.parametrize(
    "line",
    ["If ready Then n = 1: Erase values", "If ready Then n = 1: ReDim values(0 To 1)"],
)
def test_an_erase_or_redim_after_the_colon_is_conditional(line: str) -> None:
    prefix = "    ReDim values(0 To 1)\n" if "Erase" in line else ""
    source = (
        "Public Sub T(ByVal ready As Boolean)\n    Dim values() As Long\n    Dim n As Long\n"
        f"{prefix}    {line}\n    Debug.Print values(0)\nEnd Sub\n"
    )
    assert _codes(source, "unallocated-dynamic-array-access") == []


def test_a_comma_touching_the_callee_opens_its_argument_list() -> None:
    # The VBE reads `Needs, 2` as `Needs , 2` and refuses both.
    source = (
        "Sub Main()\n    Needs , 2\n    Needs, 2\nEnd Sub\n"
        "Sub Needs(ByVal x As Long, ByVal y As Long)\nEnd Sub\n"
    )
    found = [d for d in analyze_module(source) if d.code == "argument-count"]
    assert [source[d.span.start : d.span.end] for d in found] == [",", ","]
    assert all("'x' is required" in d.message for d in found)
    optional = (
        "Sub Main()\n    Two, 2\nEnd Sub\n"
        "Sub Two(Optional ByVal x As Long, Optional ByVal y As Long)\nEnd Sub\n"
    )
    assert _codes(optional, "argument-count") == []


@pytest.mark.parametrize(
    ("expression", "value"),
    [("2 => 1", True), ("2 =< 1", False), ("2 >< 1", True), ("2 < > 2", False), ("1 > = 1", True)],
)
def test_a_conditional_comparison_reads_every_spelling(expression: str, value: bool) -> None:
    assert evaluate_conditional_expression(expression) is value
