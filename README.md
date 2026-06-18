# pyVBAanalysis

Pure-Python static analysis for VBA, ported from the
[XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode) analyzer.

pyVBAanalysis takes VBA source text and returns diagnostics, plus the analysis
model behind them (tokens, AST, symbols, types). The goal is full static-analysis
parity with XLIDE under the same no-false-positive discipline: a diagnostic ships
only when it is provably correct (backed by MS-VBAL, the Excel/VBE oracle, or
deterministic metadata), and anything unknown or ambiguous stays quiet.

This is an independent, pure-analysis library. The only runtime dependency is
[pyOpenVBA](https://pypi.org/project/pyOpenVBA/), used to read VBA modules
directly from Excel, Word, and PowerPoint files; pyOpenVBA is itself pure Python
with no transitive dependencies, so the whole runtime tree is pure Python.
Everything else is the standard library. There is no Excel COM, and the library
does not run the oracle. XLIDE owns the Excel/VBE oracle and the evidence
pipeline; pyVBAanalysis consumes the emitted evidence (the oracle case corpus and
the provenance audit) as porting spec and as test fixtures.

## Status

Planning and scaffolding. The build plan, parity inventory, Python module layout,
port order, and the XLIDE sync strategy live in [agent.md](agent.md). The agent
operating policy is in
[docs/agentic_ai_programming_best_practices.md](docs/agentic_ai_programming_best_practices.md).

## Development

```
pip install -e ".[dev]"
pytest
ruff check .
mypy pyvbaanalysis
```

## License

Not yet chosen.
