# Contributing

pyVBAanalysis is a faithful Python port of the
[XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode) VBA analyzer. The
overriding rule is the no-false-positive discipline: a diagnostic ships only when
it is provably correct. When in doubt, stay quiet.

## Development setup

Python 3.10 or later.

```
pip install -e ".[dev]"
```

The `dev` extra adds the test, lint, and type tools. pyOpenVBA (the one runtime
dependency, used to read VBA out of Office files) is installed with the base
package, so the container reader is exercised by the test suite.

## Local checks

Run all three before sending a change. They are the same gate CI enforces:

```
pytest
ruff check .
mypy pyvbaanalysis
```

`mypy` runs in `--strict` mode. The suite is fast (a few seconds); run it often.

## The XLIDE port model

This repository ports XLIDE's TypeScript analyzer. XLIDE owns the language and
host knowledge; pyVBAanalysis reproduces its behavior in Python.

* The XLIDE source is expected as a sibling checkout at `../xlide_vscode`, to read
  while porting.
* XLIDE also owns the oracle: the Excel/VBE evidence corpus and the provenance
  audit. pyVBAanalysis consumes the emitted evidence verbatim as both porting spec
  and test fixtures. Do not reimplement the oracle here.
* The vendored data lives in `pyvbaanalysis/data/` (the host object models, the
  runtime tables, the oracle cases, the rule metadata, and a manifest with
  checksums). It is never hand-edited. A sync regenerates all of it from one XLIDE
  commit:

  ```
  python tools/vendor_data.py --ref vX.Y.Z
  ```

  The commit is pinned first. `tools/pin_analyzer.py` copies the sibling
  repository's history into `artifacts/analyzer-pin/<commit>` (not in source
  control) and checks that commit out, so neither uncommitted work in the sibling
  checkout nor the branch it happens to be on reaches the data. The
  `tools/extract_*.mjs` scripts then read the pin through `XLIDE_ROOT`, and stop
  rather than fall back to the sibling when it does not name an XLIDE checkout. The
  manifest records both the version label and the full commit (`xlideVersion`,
  `xlideCommit`).

The build plan, parity inventory, and module-by-module port map are in
[agent.md](agent.md).

## Checking a sync against upstream

The port's own tests encode the behavior it had, so they stay green while upstream
widens a rule the port never took. A sync is checked by running both analyzers on
the same inputs, upstream from the pin of the vendored commit:

```
python tools/differential/harness.py record
python tools/differential/harness.py replay
python tools/differential/harness.py corpus
python tools/differential/harness.py cases artifacts/differential/cases/mine.json PATH ...
python tools/differential/harness.py projects artifacts/differential/cases/mine.json
```

* `record` runs upstream's own test suite with a recorder appended to its
  `analyzeModule` and `ProjectIndex`, writing every call to
  `artifacts/differential/<commit>/calls.jsonl`. `replay` runs each call through
  the port, rebuilding the project index behind a call made with project context,
  and compares the diagnostics by code, span and message. It finds the gaps
  nothing else exercises: whole checks never ported, fixes older than the last
  sync point.
* `corpus` runs the oracle corpus through both analyzers, each module standalone
  and with its case's modules as a project.
* `cases` turns Office files, and folders searched for them and for exported
  `.bas`/`.cls`/`.frm` modules, into project cases. `--markdown` adds every VBA
  block of upstream's syntax corpus, and `--host` names a host for the cases no
  file implies one for. `projects` runs a cases file through both analyzers.

`replay`, `corpus` and `projects` exit 1 when anything differs and print what only
one side reports. Where the port reports more, the port is wrong. Where upstream
reports more it may be an upstream false positive, so check before porting it.

The upstream halves (`tools/differential/upstream/*.mjs`) read the pin through
`tools/xlide_source.mjs`, as the extractors do, and need nothing installed. Only
`record` runs upstream's vitest suite, which needs the sibling checkout's
dependencies installed. The pin gets a `node_modules` folder of its own with one
link per package, so the caches Vite writes stay in the pin, and the run fails if
the sibling's folder changed. The recorder is removed with `git checkout` when the
run ends. `harness.py unpatch` restores a pin a run left patched, and
`tools/pin_analyzer.py` will not reuse a pin with local changes.

## Adding or changing a rule

* Port the rule from its XLIDE source; match the behavior, not just the shape.
* A rule ships only with oracle backing. Validate it against the vendored corpus
  and add direct tests for the positive and the no-false-positive control cases.
* The diagnostic rule registry order is a contract: it is the diagnostic
  output order, and it must remain a faithful subsequence of XLIDE's `registry.ts`.
  Place a new entry at its XLIDE position; do not reorder existing entries.
* Diagnostic codes and their metadata (default severity, category, evidence basis,
  spec reference) live in the rule metadata. Regenerate the catalogue after a
  metadata change:

  ```
  python tools/generate_diagnostics_catalogue.py
  ```

## Style

* Write plain ASCII. No em dashes or other non-ASCII in code, comments,
  docstrings, or commit messages; use a comma, colon, period, or parentheses.
  Test fixtures that deliberately exercise Unicode are the only exception.
* Avoid AI tells and marketing language. Prefer clear, direct prose.
* Keep changes small and focused, with names and structure that match the
  surrounding code.
* Update the docs when behavior, setup, the API, or the architecture changes.
* Report status honestly: "done" means the checks above were run and passed.

## Continuous integration

`.github/workflows/ci.yml` runs the lint, type, and test gate on every push to
`main` and every pull request, across Python 3.10 to 3.13. Dependabot
(`.github/dependabot.yml`) opens weekly pull requests to update the GitHub Actions
and the Python tooling.

## Releasing

Publishing is automated by `.github/workflows/publish.yml`, which builds the sdist
and wheel and uploads them to PyPI through Trusted Publishing (OIDC, no API
tokens) when a GitHub Release is published.

One-time setup before the first release:

1. On PyPI, go to Account -> Publishing and add a pending publisher with project
   name `pyvbaanalysis`, owner `WilliamSmithEdward`, repository `pyVBAanalysis`,
   workflow `publish.yml`, and environment `pypi`.
2. On GitHub, go to repo Settings -> Environments and create an environment named
   `pypi`. Optionally add required reviewers so a publish needs human approval.

To cut a release:

1. Bump `__version__` in `pyvbaanalysis/__init__.py` (the single source of the
   version) and update `CHANGELOG.md`. Commit.
2. Tag and push: `git tag vX.Y.Z` then `git push origin vX.Y.Z`.
3. Publish a GitHub Release for that tag (for example
   `gh release create vX.Y.Z --notes-file <notes>`). Publishing the release runs
   `publish.yml`, which re-runs the gate, builds, and uploads to PyPI.
