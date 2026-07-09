# Changelog

All notable changes to pyVBAanalysis are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project aims to
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html): a minor version
per milestone.

## 1.2.0 - 2026-07-08

Brings the port up to date with upstream XLIDE v2.5.12 (the previous release
mirrored v2.5.4). The vendored data package is re-pinned from XLIDE v2.5.0 to
v2.5.12: 412 oracle cases (was 397), 120 audited codes (was 117), and a
117-rule catalogue (was 115).

### Added

Three new diagnostic codes (XLIDE v2.5.5-v2.5.6):

* `mismatched-end-keyword` (warning): a procedure closed with the wrong `End`
  keyword (e.g. `Property Get ... End Function`) still compiles - the VBE
  accepts `End Sub`/`End Function`/`End Property` interchangeably - so the
  parser now treats the procedure as closed and reports a style warning
  anchored on the opener, instead of the previous missing-closer plus
  unmatched-closer error pair.
* `call-statement-multi-arg-parens` (error): a standalone (non-`Call`)
  statement that wraps two or more arguments in parentheses
  (`mySub2("a", "b")`) is the VBE "Expected: =" compile error. Scoped to
  callees that resolve to known procedures; single-argument ByVal grouping and
  object member calls stay silent.
* `if-reserved-keyword-in-condition` (error): a reserved If-control keyword
  (`If`/`Then`/`Else`/`ElseIf`) inside a block-If condition (`If If True Then`,
  `If True Then Then`) is a VBE Syntax error. Only keyword tokens match, so
  identifiers containing those words are never flagged.

New analysis behavior:

* Juxtaposed value expressions in an assignment RHS (`n = 1 n 1`) are reported
  as `invalid-expression-syntax` ("expected end of statement"), mirroring
  XLIDE v2.5.9.
* `#If` evaluation now handles the relational operators (`<`, `>`, `<=`, `>=`)
  and hex/octal literals (`&HFF`, `&O17`).
* Indexed collection accessors resolve to their element type regardless of the
  accessor's member kind (`ws.ChartObjects(1).Chart` now resolves); the
  explicit element accessors `Item`/`_Default`/`Add` are no longer re-indexed
  (a collection-of-collections such as `SparklineGroups.Item(1)` stops
  over-resolving); a mixed-element collection (`Sheets(...)`) resolves through
  its union surface; and empty parentheses count as a call, not indexing
  (XLIDE v2.5.10).

### Fixed

Mirrors the XLIDE v2.5.11 fixes for five false-positive families found in a
real-world library:

* The `Access` grammar word of an `Open ... For mode Access Read/Write` clause
  is no longer reported as `undeclared-variable`.
* A Byte array assigned to a String scalar (`s = bytes`, the documented VBA
  encoding conversion) is no longer reported as `array-assignment-to-scalar`.
* A `ReDim` inside a single-line `If cond Then ReDim a(...)` (and its `Else`
  arm) is recognized as an allocation instead of an unallocated access.
* The type name after `As` in `ReDim x(...) As TypeName` is no longer reported
  as `undeclared-variable`.
* A parenless call whose first argument is a parenthesized group
  (`AssertTrue (cond), "msg"`) now counts every argument instead of reporting
  a wrong `argument-count`.

And the XLIDE v2.5.8 adversarial-review analyzer fixes:

* One throwing rule can no longer blank a module's whole diagnostics pass; the
  engine isolates each rule and each shared walk.
* Word operators (`And`/`Or`/`Is`/`Mod`/...) are recognized in
  invalid-operator-sequence detection; the division-by-zero constant lookup no
  longer mis-matches a longer member chain (`a.Zero.Foo`); the
  runtime-argument-value rule rejects `obj.vba.Left(...)` receiver chains; a
  call with an out-of-range argument count reports one arity diagnostic
  instead of two; every `RaiseEvent` on a `:`-separated line is checked; and a
  bare `set-required` assignment target now recognizes project class types
  (`Dim a As SomeClass : a = Null`).
* Recursion-depth guards on the expression parser and the integer-constant
  evaluator keep pathological nesting within the "never throws" contract.
* Document/UserForm code names resolve as project globals in definition
  resolution, and host-global lookup is O(1).

Also mirrors the XLIDE v2.5.12 fix for a regression the v2.5.8 word-operator
widening introduced: a `Case Is > 5` comparison clause (MS-VBAL 5.4.2.10) is
grammar, not an operator run, so it is no longer reported as
`invalid-expression-syntax` (operator runs inside a Case body still are).

Also mirrors the XLIDE v2.5.12 juxtaposition fix: a `&`-suffixed integer
literal followed by a value (`s = 3000000000&"x"`) is not reported as
juxtaposed, because the VBE can read that `&` as concatenation (oracle case
`suffix_long_amp_glued_concat_accepted`).

### Development

A new drift gate (`tests/test_registry_parity.py`) statically proves every
catalogue rule has a Python emitter, so re-vendoring the data package after a
future upstream release turns CI red on exactly the rules still to be ported.

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
