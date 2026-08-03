"""One native-language VBA sample per supported code page, run end to end.

XLIDE shipped this matrix after its issue #6: an ASCII-only assumption in the
analyzer made a Cyrillic-named `Sub` invisible to the structural pass, so its
`End Sub` reported "no matching Sub" on perfectly valid code. The bug class is
"some regex or predicate in the pipeline assumes ASCII", and it can hide in the
lexer, the parser, the symbol builder, or any rule.

The scope here is deliberate. pyVBAanalysis analyzes VBA *source text*, so this
matrix drives the analyzer with native-language identifiers, string literals and
comments and asserts the whole stack stays correct and quiet. Turning workbook
bytes into that text is pyOpenVBA's job (a project's code page decides what its
bytes mean, and pyOpenVBA 3.4.0 is the floor precisely because it resolves code
pages correctly) and is covered by that project's own matrix, so it is not
re-tested here.

Each row asserts four things:

1. The token stream round-trips: raw text plus trivia reconstructs the source
   byte for byte. Round-trippability is the lexer's acceptance gate (agent.md
   M1), and it is what catches a scanner that mis-measures a multi-byte or
   astral character.
2. The parser finds the procedure under its native name, with no parse
   diagnostics.
3. The full analyzer is silent on the sample, which is valid VBA.
4. The same source analyzed with whole-project context stays silent, so the
   cross-module rules (undeclared-variable, unknown-call, member-not-found)
   do not fire on non-ASCII names either.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_project
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.parser.nodes import ProcedureNode
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind

# (label, procedure name, local variable name, string literal / comment text).
# The identifiers mirror the code pages XLIDE's matrix covers; the analyzer sees
# decoded text, so the code page itself is not a parameter here - the point is
# that every one of these scripts is exercised as an identifier and as data.
LANGUAGE_MATRIX: list[tuple[str, str, str, str]] = [
    ("Thai (cp874)", "ทดสอบ", "ผล", "ทดสอบภาษาไทย"),
    ("Japanese (cp932)", "テスト", "値", "テスト用モジュール"),
    ("Chinese Simplified (cp936)", "测试", "值", "中文测试模块"),
    ("Korean (cp949)", "테스트", "값", "한국어 테스트"),
    ("Chinese Traditional (cp950)", "測試", "值", "繁體中文測試"),
    ("Central European (cp1250)", "Zkouška", "hodnota", "Příliš žluťoučký kůň"),
    ("Cyrillic (cp1251)", "Проверка", "значение", "Проверка русского текста"),
    ("Western European (cp1252)", "Prüfung", "wert", "déjà vu œuvre Straße"),
    ("Greek (cp1253)", "Δοκιμή", "τιμή", "Δοκιμή ελληνικού κειμένου"),
    ("Turkish (cp1254)", "Deneme", "değer", "Türkçe deneme ğüşiöç İı"),
    ("Hebrew (cp1255)", "בדיקה", "ערך", "בדיקת עברית"),
    ("Arabic (cp1256)", "اختبار", "قيمة", "اختبار العربية"),
    ("Baltic (cp1257)", "Bandymas", "reikšmė", "Lietuviškas tekstas ąčęėįšųū"),
    ("Vietnamese (cp1258)", "ThửNghiệm", "giáTrị", "Tiếng Việt thử nghiệm"),
    ("KOI8-R (cp20866)", "Тест", "переменная", "Тест КОИ-8"),
    ("KOI8-U (cp21866)", "ТестУкр", "змінна", "Тест української ґї"),
    ("ISO-8859-2 (cp28592)", "Zazil", "zmienna", "Zażółć gęślą jaźń"),
    ("UTF-8 (cp65001)", "Смешанный", "混合", "любой текст 中文 déjà ทดสอบ"),
]

_IDS = [row[0] for row in LANGUAGE_MATRIX]


def _module_source(proc: str, local: str, text: str) -> str:
    return (
        'Attribute VB_Name = "LangModule"\r\n'
        "Option Explicit\r\n"
        "\r\n"
        f"' {text}\r\n"
        f"Public Sub {proc}()\r\n"
        f"    Dim {local} As String\r\n"
        f'    {local} = "{text}"\r\n'
        f"    Debug.Print {local}\r\n"
        "End Sub\r\n"
    )


def _reconstruct(source: str) -> str:
    """Rebuild the source from the token stream, trivia included."""
    out: list[str] = []
    for token in tokenize(source):
        for trivia in token.leading_trivia:
            out.append(trivia.text)
        out.append(token.raw_text)
        for trivia in token.trailing_trivia:
            out.append(trivia.text)
    return "".join(out)


@pytest.mark.parametrize(("label", "proc", "local", "text"), LANGUAGE_MATRIX, ids=_IDS)
def test_tokens_round_trip(label: str, proc: str, local: str, text: str) -> None:
    source = _module_source(proc, local, text)
    assert _reconstruct(source) == source


@pytest.mark.parametrize(("label", "proc", "local", "text"), LANGUAGE_MATRIX, ids=_IDS)
def test_parses_under_the_native_name(label: str, proc: str, local: str, text: str) -> None:
    source = _module_source(proc, local, text)
    module = parse_module(source)
    procedures = [m for m in module.members if isinstance(m, ProcedureNode)]
    assert [p.name for p in procedures] == [proc]
    assert [d.message for d in module.diagnostics] == []


@pytest.mark.parametrize(("label", "proc", "local", "text"), LANGUAGE_MATRIX, ids=_IDS)
def test_analyzer_is_silent(label: str, proc: str, local: str, text: str) -> None:
    source = _module_source(proc, local, text)
    diagnostics = [f"{d.code}: {d.message}" for d in analyze_module(source)]
    assert diagnostics == []


@pytest.mark.parametrize(("label", "proc", "local", "text"), LANGUAGE_MATRIX, ids=_IDS)
def test_analyzer_is_silent_with_project_context(
    label: str, proc: str, local: str, text: str
) -> None:
    # A whole-project view turns on the cross-module rules, which resolve names
    # through the symbol graph - the layer most likely to mishandle non-ASCII.
    source = _module_source(proc, local, text)
    results = analyze_project([ModuleInput("LangModule", ModuleSymbolKind.STANDARD, source)])
    diagnostics = [f"{d.code}: {d.message}" for d in results["LangModule"]]
    assert diagnostics == []


def test_matrix_covers_the_supported_pages() -> None:
    # A floor, so trimming the matrix is a deliberate edit rather than an
    # accident: every row above is a distinct script family.
    assert len(LANGUAGE_MATRIX) >= 18
    assert len({row[1] for row in LANGUAGE_MATRIX}) == len(LANGUAGE_MATRIX)


def test_matrix_identifiers_avoid_combining_marks() -> None:
    # Keeps the matrix honest about its own scope: the identifier columns use
    # scripts whose letters stand alone, because a combining mark in an
    # identifier is a known gap (see the xfail below). Sample TEXT is not
    # restricted - several rows carry marks there, which is the point.
    import unicodedata

    for label, proc, local, _text in LANGUAGE_MATRIX:
        for name in (proc, local):
            marks = [c for c in name if unicodedata.category(c).startswith("M")]
            assert not marks, f"{label}: identifier {name!r} carries a combining mark"


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Known gap: the lexer treats an identifier character as a letter via "
        "str.isalpha(), which is False for combining marks (category Mn/Mc). A "
        "Thai identifier like 'kha + mai ek + sara aa' therefore splits and the "
        "fragment is reported undeclared. Whether the VBE accepts such an "
        "identifier at all needs Excel/VBE oracle evidence before the lexer "
        "changes - inventing the rule would violate the no-false-positive "
        "discipline in either direction. Flip this test when that evidence lands."
    ),
)
def test_combining_mark_identifier_is_not_split() -> None:
    source = _module_source("ทดสอบ", "ค่า", "ทดสอบภาษาไทย")
    results = analyze_project([ModuleInput("LangModule", ModuleSymbolKind.STANDARD, source)])
    assert [f"{d.code}: {d.message}" for d in results["LangModule"]] == []


def test_non_ascii_names_still_resolve_across_modules() -> None:
    # The cross-module path: one module calls another's natively-named Sub. If
    # the symbol graph folded non-ASCII names wrongly this reports unknown-call.
    callee = (
        'Attribute VB_Name = "Библиотека"\r\n'
        "Option Explicit\r\n"
        "\r\n"
        "Public Sub Приветствие()\r\n"
        "End Sub\r\n"
    )
    caller = (
        'Attribute VB_Name = "Вызов"\r\n'
        "Option Explicit\r\n"
        "\r\n"
        "Public Sub Запуск()\r\n"
        "    Приветствие\r\n"
        "End Sub\r\n"
    )
    results = analyze_project(
        [
            ModuleInput("Библиотека", ModuleSymbolKind.STANDARD, callee),
            ModuleInput("Вызов", ModuleSymbolKind.STANDARD, caller),
        ]
    )
    assert [f"{d.code}: {d.message}" for d in results["Вызов"]] == []
    assert [f"{d.code}: {d.message}" for d in results["Библиотека"]] == []
