# Changelog

All notable changes to pyVBAanalysis are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project aims to
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html): a minor version
per milestone.

## 1.1.1 - 2026-06-20

### Fixed

Mirrors upstream XLIDE v2.5.x fixes that eliminate false positives on real-world
workbooks (validated against the fastjson and stdVBA workbooks):

* Recognize the hidden VBA intrinsics `VarPtr`/`StrPtr`/`ObjPtr` and the byte-string
  family (`LeftB`/`RightB`/`MidB`/`InStrB`/`AscB`/`ChrB` and the `$` variants), plus the
  `vbLongLong` constant, so they are no longer reported as `undeclared-variable`.
* A qualified `ReDim` target (`ReDim x.arr(...)`) resizes a member array and is no longer
  misreported as `scalar-redim` on the container variable.
* `Exit Function` / `Exit Sub` inside a `Property Get` are accepted (the VBE allows
  them), so they no longer raise `exit-wrong-proc`.
* The mandatory value parameter of a `Property Let`/`Set` may follow an `Optional` index
  parameter without a `required-param-after-optional` error.
* Default the `TWINBASIC` compiler constant to False so twinBASIC-only `#If` branches are
  inactive, and compare boolean `#Const` values by their VBA numeric form (`-1`/`0`).
* Harden token-name handling against an empty token, matching the upstream null-guard.

## 1.1.0 - 2026-06-20

### Added

* Inline suppression: `'@pyvba-ignore`, `'@pyvba-ignore-next-line`, and
  `'@pyvba-ignore-file` comment directives suppress diagnostics from within the source
  (optional comma-separated code list, case-insensitive, with a `-- reason` trailer). A
  malformed directive is reported as `analysis-suppression-directive`. A new
  `inline_suppression` option and a `--no-inline-suppression` CLI flag turn it off for
  an audit run.
* A `whole_project` flag on `analyze_project`, `analyze_loose_file` (default False for a
  single file), and `analyze_loose_files`, plus a `--partial-project` CLI flag and
  automatic partial treatment of a single targeted file.
* Usage-guide sections for inline suppression, "Whole project vs a single file", and
  "Use in CI".

### Fixed

* Workbook reader: class modules read out of a workbook were misclassified as document
  modules (they carry a `VB_Base` line like documents, but with the generic VBA class
  base GUID). Classification now keys on the GUID, so `New SomeClass` for a workbook
  class is no longer reported as `invalid-new-type-name`.
* Single-file analysis no longer emits the whole-project checks (`undeclared-variable`,
  `unknown-call`, `member-not-found`) as false positives: a rule that needs every module
  is skipped when the analyzed set is not the complete project, since a symbol declared
  in an unseen module is indistinguishable from an undefined one.

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
