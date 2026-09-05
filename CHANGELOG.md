# Changelog

All notable changes to pyVBAanalysis are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project aims to
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html): a minor version
per milestone.

## 2.1.1 - 2026-09-05

Follows XLIDE 6.2.0. The vendored data is byte-identical to 6.1.2, so only the
manifest's version moves; the analyzer change is the one fix below.

### Fixed

* A module's own procedures now shadow the host's globals when used as a
  receiver. A module VARIABLE named `rows` already resolved from its
  declaration, but `Public Property Get rows() As Widget` fell through to
  Excel's global `Rows`, so `rows.Where(p)` was measured against `Excel.Range`.
  The names that collide are the ones every workbook uses: rows, columns,
  cells, selection, names, sheets, application. A Function or Property Get now
  yields its return type; a Sub, or a Property with only Let/Set, yields nothing
  readable but still shadows the global rather than letting the host answer
  (XLIDE issue #68, which this port reported upstream and which 6.2.0 fixed).

  This port never emitted the false positive, because it resolved such a
  receiver to nothing at all rather than to the wrong type. The fix replaces
  that silence with a correct binding, so a receiver that used to go unchecked
  is now checked: with project context, a genuinely absent member on one of
  these names reports `member-not-found` where it previously passed. Analyzed
  standalone, without project context, behaviour is unchanged.

### Verified

* Differential against the upstream 6.2.0 analyzer: 418 of 418 oracle cases
  identical in both directions, and identical on a real workbook's modules.
  16 real workbooks stay silent.

## 2.1.0 - 2026-09-05

Sync to XLIDE 6.1.2, the analyzer's first re-pin since 3.1.4. Verified by a
differential against the upstream analyzer over the whole oracle corpus: 418 of
418 cases identical, in both directions.

### Added

* `ambiguousProjectProcedure`: VBA is content for two modules to export the same
  public procedure name, but it refuses to compile an UNQUALIFIED call to that
  name from a module declaring neither. The finding sits at the call site, not
  the declarations, because a project that exports a name twice and always
  qualifies its calls is legal VBA and common. Silent when the call is
  qualified, when the calling module declares the name itself, when a local or
  parameter shadows it, or when only one module exports it.
* `ConditionalActivityTracker.mutually_exclusive` and `in_same_branch`, the
  primitive the arm-aware rules below need, plus
  `ProjectIndex.implemented_interface_names()` and the
  `AnalyzeModuleOptions.implemented_interfaces` option it feeds.

### Fixed

* Rules no longer pair declarations from mutually exclusive `#If` arms. Only one
  arm of a chain is ever built, so two declarations in different arms are
  alternatives rather than duplicates. With a compiler constant the analyzer
  cannot evaluate, the ordinary `#If VBA7 / #Else` idiom reported
  `duplicate-procedure` and `duplicate-module-variable` on legal code. The same
  reasoning now covers duplicate declarations and labels, undefined labels,
  `For`/`Next` pairing, `Else` branch order, `Option` placement, and `Implements`
  placement (XLIDE issue #58).
* A `#Const` directive may precede `Option Explicit`. A conditional-compilation
  directive is not a declaration, and the live VBE compiles it there
  (oracle case `const_directive_before_option_explicit_compile`).
* Three gaps in return-assignment detection, each of which the widened rule below
  would otherwise have turned into a false positive:
  assigning a field of the returned value (`MsToSystemTime.wYear = ...`) is a
  return assignment; so is an assignment inside a single-line `If`, whose
  branches the statement walk never entered; and a name that SPELLS a keyword is
  still an assignment target, so `Function Read()` assigning `Read = True` was
  read as never assigning its return. Measured on one real workbook these three
  accounted for nine false findings.

### Changed

* `missingReturnAssignment` now covers every Function and Property Get, not only
  untyped ones, so a typed Function that never assigns its return reports. An
  empty member of a module some other module declares with `Implements` stays
  silent: that is a contract for an implementer to fill in, not unfinished code.
  A body whose work is to raise stays silent too. **This reports on code that was
  previously quiet**; the rule is warning severity and can be turned off through
  `severity_overrides`.
* The vendored host models grew with upstream's move to the whole documented
  object model: Excel 229 to 651 types, Word 364 to 627, PowerPoint 201 to 464,
  Access 188 to 451. The wheel grows from 0.94 MB to 1.84 MB; the models still
  load lazily, so import cost is unchanged.
* Data re-pinned to XLIDE 6.1.2: 418 oracle cases (was 415), 122 audited codes
  (was 121), 119 rules (was 118).

### Known limits

* Upstream's form-control member work is not ported, and the oracle corpus has no
  UserForm designer cases, so that area is untested here rather than verified.

## 2.0.0 - 2026-08-19

Multi-host analysis: VBA is now measured against the object model of the Office
host it actually belongs to. Ports the XLIDE host-model seam and the Word,
PowerPoint and Access models (xlide_vscode 790e6ea and e56098b, its issues #24
and #25).

### Added

* A host seam on the analysis entry points. `analyze_project` and
  `analyze_module_options_for` take a `host` token (`"excel"`, `"word"`,
  `"powerpoint"`, `"access"`, ...), and `AnalyzeModuleOptions` carries `host`
  alongside the existing `host_model`. The semantics are deliberately
  asymmetric: absent means Excel, so every existing caller is unchanged, while a
  NAMED host with no model asserts no host knowledge rather than falling back to
  Excel's. An explicit `host_model` still outranks the token.
* Word, PowerPoint and Access object models, vendored as
  `data/{word,powerpoint,access}_host_model.json` and extracted mechanically by
  `tools/extract_host_model.mjs` from the generated XLIDE host modules, the same
  pipeline that has always produced the Excel model. 364, 201 and 188
  member-bearing types; 3,742, 1,480 and 1,116 enum constants. Every type is
  non-exhaustive by construction, so absence never becomes a finding.
* `analyze_office_file(path)` and `read_office_modules(path)` read any container
  pyOpenVBA opens: Excel (`.xlsm`, `.xlsb`, `.xlam`, `.xls`), Word (`.docm`,
  `.dotm`, `.doc`), PowerPoint (`.pptm`, `.potm`) and Access (`.accdb`, `.mdb`,
  read-only). The extension selects the host, so a Word document is analyzed
  against Word's model without the caller saying anything. The CLI accepts them
  all on the same footing.
* `host_token_for_file_name` and `host_object_model_for_token` on the public API.

### Fixed

* Word VBA no longer false-positives against Excel's object model. Analyzing a
  Word module as a project previously reported four findings on legal code:
  `ActiveDocument` twice and `wdOrientPortrait` as undeclared variables, and
  `Selection.TypeText` as `member-not-found` against `Excel.Range`. Under the
  Word host it reports none.
* Host metadata is per-model throughout, closing the Excel defaults that
  remained under the seam. Application-member injection is keyed by model
  (`Volatile` is a known bare call under Excel and unknown under Word), and the
  four rules that consulted host metadata with a bare Excel default now take the
  caller's model: undeclared variables, unknown call, ambiguous enum references,
  and late-bound Friend members. `Me` types host-correctly in a document module:
  `ThisWorkbook` is `Excel.Workbook` under Excel, `ThisDocument` is
  `Word.Document` under Word, and nothing is asserted anywhere else.
* Word's `ThisDocument` classifies as a document module. It declares
  `VB_Base = "1Normal.ThisDocument"`, naming no CLSID, so GUID matching missed
  it; the host-generic signature is the `VB_PredeclaredId` + `VB_Exposed` pair
  that Office gives every document module. A `.bas` extension or the container's
  own standard-module flag still outranks that signature, since either states
  outright that the module is standard. Differentialled across 116 modules in 16
  real workbooks: nothing reclassified.
* A named host with no object model no longer reports what it cannot know. An
  empty model is not uniformly quieter than Excel's: member lookups go silent
  because no type resolves, but the rules that ask whether a bare name is legal
  would answer no for every global the host injects, turning an Outlook
  project's own surface into a wall of `undeclared-variable` findings. The four
  rules that need host knowledge now stay silent under an unmodelled host, for
  the same reason they already do on a partial project view.
* The host-model memos no longer key on a bare `id(model)`. Three caches
  (`_host_model_index`, `_host_constant_index`, and the host-member name set)
  used a plain `dict[int, ...]` that kept no reference to the model, so entries
  accumulated one per model object ever seen and a collected model's id could be
  recycled by a later one, serving one model's index for another. They now use
  the repo's `IdentityLru`, which holds its keys alive and stays bounded. Latent
  before this release, since only the permanently-live Excel singleton was ever
  passed; reachable now that callers choose models.

### Changed

* `analyze_workbook` and `read_workbook_modules` keep their Excel-only contract
  unchanged. Handed a container the generic reader could open, they now name
  `analyze_office_file` in the error rather than only rejecting the extension.
* README, the usage guide, the API reference and agent.md describe the four
  hosts; agent.md's pyOpenVBA floor was stale at 3.0.1 and now reads 3.4.0.

### Known limits

* Legacy `.ppt` is not readable. pyOpenVBA 3.4.0 lists it but reads it as a
  plain CFB, while the VBA project lives in a zlib-compressed CFB inside an
  `ExOleObjStg` record, so every open fails on the missing `dir` stream. The
  extension is rejected up front rather than failing later with a parse error
  that reads like file corruption. Reported as pyOpenVBA issue #17.
* Access is read-only, following pyOpenVBA: Access executes compiled p-code, so
  a source write would silently change nothing.

## 1.4.2 - 2026-08-03

### Fixed

* Identifier continuation now accepts every Unicode mark category, matching
  XLIDE. The 1.4.1 fix listed only `Mn` and `Mc`, while XLIDE's equivalent fix
  (its issue #8) widened its patterns to `\p{M}`, which also covers enclosing
  marks (`Me`). An identifier carrying one still split here. agent.md makes the
  XLIDE TypeScript the executable spec for this port, and widening what
  continues a name can only remove false positives, so the predicate now
  accepts any category beginning with `M`.

  Enclosing marks in VBA identifiers are vanishingly rare, which is why nothing
  caught it earlier: the oracle probe covered `Mn`, and the language matrix
  covers `Mn` and `Mc`. It surfaced from verifying XLIDE's fix against this
  one, side by side.

## 1.4.1 - 2026-08-03

### Fixed

* An identifier containing a Unicode combining mark is no longer split by the
  lexer. Scripts like Thai write one letter as a base plus a tone mark and/or
  vowel sign, and `str.isalpha()` is False for those marks (categories Mn and
  Mc), so `Dim <kho + mai ek + sara aa>` was lexed as three tokens and the tail
  reported as `undeclared-variable`: a false positive on valid VBA.

  This is now VBE-oracle verified rather than assumed. A cp874 project
  declaring that identifier compiles and runs clean in real Excel, with a
  mark-free Thai control alongside it, both reporting no compile dialog.
  `tools/oracle/build_combining_mark_probe.py` rebuilds the probes so the
  finding is reproducible; the corpus entry belongs upstream in XLIDE, which
  owns the evidence pipeline.

  The 1.4.0 language matrix recorded this as a strict xfail pending exactly
  this evidence. It is now an ordinary passing test, and the Thai row uses the
  combining-mark identifier again so the regression cannot return unnoticed.

## 1.4.0 - 2026-08-03

### Changed

* The pyOpenVBA floor is now 3.4.0 (was 3.0.1). That release fixes code-page
  resolution, so module text in a non-Latin project (Cyrillic, Greek, the
  double-byte CJK pages) no longer decodes as mojibake the analyzer would read
  as identifiers. Reads of Latin-1 projects are byte-identical either way, and
  the whole test suite passes unchanged on the new version.

### Performance

A second profiling pass on the same 26,721-line class module, again with
identical diagnostics. `analyze_module` drops from 3.4s to 2.5s and the
containing workbook through the CLI from 6.0s to 4.25s; cumulatively since
1.3.0 that module went from 27.2s to 2.5s. Full-module lexes per pass fell
from 38,526 to about 4,100, and the profile is now flat, with no entry above
roughly 5 percent.

* Three paths still re-lexed text the shared token stream already covers: the
  procedure-label scan (33k slice lexes per pass), the TypeOf-operand check
  (which re-lexed the entire module in one call), and the declaration
  assignment-offset helper. All three now ride the per-pass cache.
* Two lexer hot loops scan runs instead of calling a predicate per character:
  identifiers advance by an ASCII-run regex, with any non-ASCII continuation
  still checked by the exact predicate so the stopping point is unchanged, and
  whitespace trivia matches a run built from the same character set the
  predicate uses, with a fast path for the common no-trivia case.

### Added

* A CI language matrix, mirroring the one XLIDE added after its issue #6. One
  native-language sample per supported code page - Thai, Japanese, Simplified
  and Traditional Chinese, Korean, Central European, Cyrillic, Western
  European, Greek, Turkish, Hebrew, Arabic, Baltic, Vietnamese, KOI8-R/U,
  ISO-8859-2 and UTF-8 - drives the analyzer with native identifiers, string
  literals and comments, asserting the token stream round-trips exactly, the
  procedure parses under its native name, and both the module and whole-project
  passes stay silent. It runs on Linux and Windows, since the failure mode it
  guards is encoding-shaped. A cross-module case covers a natively-named Sub
  called from another module.

  The matrix also pins a known gap it uncovered: an identifier containing a
  combining mark (Thai, for instance) is split by the lexer, because character
  membership is decided with `str.isalpha()`, which is false for marks. That
  currently yields a false `undeclared-variable` on the fragment. Whether the
  VBE accepts such an identifier at all needs oracle evidence before the lexer
  changes, so the case is recorded as a strict xfail that will fail loudly if
  the behavior ever changes.

## 1.3.2 - 2026-08-03

### Internal

House style (agent.md UM-07) is plain ASCII, but em dashes had crept into four
test docstrings and into the prose of the AI-smells field guide - a document
whose own first rule is "Default to ASCII". Both are now consistent with the
rule they document.

The field guide keeps every character it is actually about: the rule that names
em dashes, en dashes, the ellipsis character and decorative emoji; the labeled
"Smells:" examples; the quoted bad patterns; and the quick-reference rows
listing them. Accented characters in cited researchers' names are also
preserved, since stripping them would misspell real people.

No analyzer behavior changes; the wheel is unaffected.

## 1.3.1 - 2026-08-02

### Performance

A profiling pass driven by a real-world 26,721-line class module. Identical
diagnostics before and after (the full suite, the oracle corpus sweeps, and a
new statement-token equivalence sweep all gate the change); the module's
analysis time drops from 27.2s to 3.4s (8x), and the whole containing
workbook through the CLI from 35s to 6s.

* Procedure-invariant derivations are no longer rebuilt per rule x procedure:
  the source-name shadow scope, the module non-callable index, the visible
  identifier-name base, the type and declaration-shape environments (module
  portion cached, cloned per procedure), and the procedure-symbol lookup
  (indexed by span start) are all memoized by object identity through a small
  bounded IdentityLru (mirroring upstream's WeakMap caches).
* One lex per module: every statement's token view is now derived from the
  module's shared memoized tokenization (binary search + offset rebase, with a
  fallback to slice-lexing when a span does not align), and the parser, the
  member-completion context, the inline-suppression scan, the call-shape
  helpers, and the unstructured-flow scan all ride the same cached stream
  instead of re-lexing per statement or per pass.
* Bare identifier references resolve through a per-module name index instead
  of scanning every module-level declaration per reference; bare type names
  resolve through a per-(project types, host model) candidate index; member
  receiver prefixes slice from the previous newline token instead of copying
  the whole module prefix; and the enclosing-procedure lookup binary-searches
  an indexed span table.

## 1.3.0 - 2026-08-01

### Added

A new oracle-backed rule, ported from XLIDE (issue #5): `late-bound-friend-member`.

Friend members are not on a class's IDispatch interface, so reaching one
through a receiver whose static type is Variant or Object raises runtime error
438 - and the compiler says nothing, because it cannot know the runtime type
either. The code compiles clean and dies on the first execution that reaches
the call. Three VBE oracle cases back the rule, including the non-obvious one:
a class reading its OWN Friend member through an `Object` local fails
identically.

Two receiver shapes are recognized: a bare identifier declared `As Variant` /
`As Object` / with no type at all, and a `Collection` element (`coll(i)` or
`coll.Item(i)`), since `Collection.Item` returns Variant and so loses the
element type however strongly typed the collection's contents are - the shape
that hides the bug in practice.

Scoped for no false positives: it fires only when the member name resolves
exclusively to Friend members of exhaustive project class modules, and stays
silent when the name is also Public anywhere, exists in the host object model
or a VBA runtime object, the receiver is strongly typed, or the name is
unknown everywhere (the VBE oracle records unknown members on late-bound
receivers as compile-valid).

The vendored data package is re-pinned to XLIDE v3.1.4: 415 oracle cases,
121 audited codes, a 118-rule catalogue. The evidence files are the only
analyzer-relevant upstream change since v2.5.12.

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
