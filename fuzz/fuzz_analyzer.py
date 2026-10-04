"""Coverage-guided fuzzing of the analyzer on arbitrary VBA source.

The analyzer reads source from untrusted Office files, so every stage must
take any text:

  lexer   tokenize() never raises, and its tokens round-trip: leading trivia
          plus raw text, plus the last token's trailing trivia, give back the
          source exactly. A source of trivia alone (whitespace, line
          continuations) yields no tokens, as tokenize() documents, so there
          is nothing to round-trip.
  parser  parse_module() never raises.
  rules   every diagnostic rule runs without raising. analyze_module() hides
          an exception by returning no diagnostics, so this target calls the
          rule runner beneath it, where a crash is visible.

Input bytes are decoded as UTF-8 with replacement, so a seed is plain source.

    python fuzz/fuzz_analyzer.py <target> [libFuzzer options] [corpus dirs]
    python fuzz/fuzz_analyzer.py lexer -max_total_time=60 tests/fuzz_corpus/vba

The .github/workflows/fuzz.yml workflow runs each target from
tests/fuzz_corpus/vba. A finding becomes a seed there and a test.
"""

import sys

import atheris

with atheris.instrument_imports():
    from pyvbaanalysis import AnalyzeModuleOptions
    from pyvbaanalysis.diagnostics.analyze_module import _run_rules, with_resolved_host_model
    from pyvbaanalysis.lexer import tokenize
    from pyvbaanalysis.parser.parse_module import parse_module

OPTIONS = with_resolved_host_model(AnalyzeModuleOptions())


def source_of(data):
    return data.decode("utf-8", errors="replace")


def fuzz_lexer(data):
    source = source_of(data)
    tokens = tokenize(source)
    if not tokens:
        return
    parts = []
    for token in tokens:
        parts.extend(trivia.text for trivia in token.leading_trivia)
        parts.append(token.raw_text)
    if tokens:
        parts.extend(trivia.text for trivia in tokens[-1].trailing_trivia)
    if "".join(parts) != source:
        raise AssertionError("the token stream does not round-trip to the source")


def fuzz_parser(data):
    parse_module(source_of(data))


def _raise_internal_error(error, where):
    raise error


def fuzz_rules(data):
    _run_rules(source_of(data), OPTIONS, _raise_internal_error)


TARGETS = {"lexer": fuzz_lexer, "parser": fuzz_parser, "rules": fuzz_rules}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_analyzer.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    atheris.Setup([sys.argv[0], *sys.argv[2:]], TARGETS[sys.argv[1]])
    atheris.Fuzz()


if __name__ == "__main__":
    main()
