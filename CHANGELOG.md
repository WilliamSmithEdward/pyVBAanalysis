# Changelog

All notable changes to pyVBAanalysis are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project aims to
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html): a minor version
per milestone.

## 1.0.0 - 2026-06-20

The first public release: a pure-Python static analyzer for Excel VBA with a
no-false-positive discipline, where a diagnostic is reported only when it is
provably correct and anything unknown or ambiguous stays quiet.

### Analysis

* The complete analysis stack: lexer, parser, symbol and project index,
  conditional compilation, type inference, the Excel host object model, the
  member-completion surface, and the project-type registry.
* 85 diagnostic rules emitting a catalogue of 117 diagnostic codes, validated
  against a corpus of 397 real Excel/VBE behavior cases.

### Ingestion and entry points

* `analyze_module` for one module's source text, and `analyze_project` for a set
  of modules analyzed together with cross-module context.
* `analyze_loose_file` / `analyze_loose_files` for loose `.bas` / `.cls` / `.frm`
  export files, and `analyze_workbook` for VBA read directly out of Excel
  workbooks. `build_project_index` and `analyze_module_options_for` expose the
  per-module options for a custom pass.
* A `pyvbaanalysis.reader` package that strips the VBE export header, infers the
  module kind, and reads modules from Excel files.
* A command-line interface: `python -m pyvbaanalysis PATH ...` over loose files,
  folders, and Excel workbooks, with `--only`, `--severity`, `--select` /
  `--ignore`, `--fail-level`, `--format`, and CI-friendly exit codes.
* The headline entry points are re-exported from the package root, and the package
  ships a `py.typed` marker.

### Packaging

* MIT license and a Trusted Publishing release workflow (OIDC, no API tokens).
* One runtime dependency, pyOpenVBA, used to read VBA out of Excel workbooks and
  imported lazily.

### Documentation

* A usage guide, an API reference, a generated diagnostics catalogue, and a
  contributing guide.
