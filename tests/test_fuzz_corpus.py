"""The fuzz corpus, replayed through the three fuzz targets on every run.

tests/fuzz_corpus/vba seeds fuzz/fuzz_analyzer.py; a fuzz finding joins it as
a regression seed. Here each seed must lex into a stream that round-trips,
parse, and pass every diagnostic rule without raising.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyvbaanalysis import AnalyzeModuleOptions
from pyvbaanalysis.diagnostics.analyze_module import _run_rules, with_resolved_host_model
from pyvbaanalysis.lexer import tokenize
from pyvbaanalysis.parser.parse_module import parse_module

CORPUS = Path(__file__).parent / "fuzz_corpus" / "vba"
SEEDS = sorted(CORPUS.glob("*"))


def _source(path: Path) -> str:
    return path.read_bytes().decode("utf-8", errors="replace")


def test_the_corpus_is_present() -> None:
    assert len(SEEDS) >= 100


@pytest.mark.parametrize("path", SEEDS, ids=lambda p: p.name)
def test_a_seed_round_trips_parses_and_analyzes(path: Path) -> None:
    source = _source(path)
    tokens = tokenize(source)
    if tokens:
        parts: list[str] = []
        for token in tokens:
            parts.extend(trivia.text for trivia in token.leading_trivia)
            parts.append(token.raw_text)
        parts.extend(trivia.text for trivia in tokens[-1].trailing_trivia)
        assert "".join(parts) == source
    parse_module(source)
    _run_rules(source, with_resolved_host_model(AnalyzeModuleOptions()))
