# pyVBAanalysis - Agent Gameplan

> Status: COMPLETE. The port shipped at v1.0.0. Milestones M0 through M10 are done
> (the full lexer, parser, symbol and project index, conditional compilation, type
> inference, host object model, member-completion surface, project-type registry,
> and all 85 registry rules), and a file/workbook ingestion layer was added on top
> (analyze_project, analyze_workbook, the loose-file reader, and the CLI). This
> document is retained as the historical build plan and the XLIDE data-sync
> reference; the milestone sections below are written in the future tense they were
> planned in, and the "first action: start M0" guidance is historical. For current
> contributor guidance see CONTRIBUTING.md, and treat
> `pyvbaanalysis/data/manifest.json` as the authoritative list of vendored data
> files (the filenames sketched below predate the shipped layout).

This file is the operating plan for building **pyVBAanalysis**: an independent,
pure-Python static analyzer for VBA, ported from the proven XLIDE TypeScript
analyzer. It is both the porting gameplan and the repo-level agent instruction
file (per RG-17 in `docs/agentic_ai_programming_best_practices.md`). Read that
best-practices doc first; this plan assumes it. Write everything in plain ASCII,
no em dashes, no AI tells (UM-07).

Grounded against the XLIDE repo at the sibling checkout `../xlide_vscode` (read on
2026-06-17). Every count and path below was read from that repo, not recalled. If
a fact here ever conflicts with the XLIDE source, the source wins; re-verify
before acting (RG-18).

---

## 1. Mission and scope

Build a library that takes VBA source text and returns diagnostics, plus the
analysis model behind them (tokens, AST, symbols, types). Full static-analysis
parity with XLIDE is the goal: every diagnostic XLIDE ships, Python ships, with
the same no-false-positive guarantee.

In scope:

- Lexer, conditional-compilation activity, parser/AST, symbols and scope, name
  resolution, type inference, dataflow, the diagnostics engine, and all 117
  diagnostic codes (85 rule functions across 17 family files).
- The Excel host metadata as a data layer (it is data, not COM), because some
  rules (member-not-found, host type narrowing) need it for full parity. It is
  optional to load: pure-VBA analysis runs without it.

Out of scope:

- The Excel/VBE oracle and any COM. XLIDE owns the oracle and the evidence
  pipeline. Python never generates ground truth; it consumes the emitted data.
- Editor features: completion UI, hover, signature help, code actions, live-edit
  span heuristics. Port only the resolver cores that diagnostics depend on (see
  Risk 7), not the editor wrapping.

Dependencies: the only runtime dependency is pyOpenVBA
(https://pypi.org/project/pyOpenVBA/, version 3.0.1, requires Python >=3.10),
used to read VBA modules directly out of Excel workbooks (the analyzer targets
Excel VBA; the shipped reader is Excel-only).
pyOpenVBA is itself pure Python with no transitive dependencies, so the whole
runtime tree stays pure Python; everything else is the standard library. The
analysis core operates on source text and never imports pyOpenVBA. Only the
reader layer (Section 3) imports it, so the core is testable without any Office
file. Dev-only tooling (pytest, ruff, mypy) is not a runtime dependency.

The hard part is already done. XLIDE is a proven reference implementation and it
carries a language-agnostic evidence corpus (397 oracle-verified cases plus a
provenance audit and an MS-VBAL verification map). That turns this from research
into a porting project with a correctness oracle at every step.

---

## 2. The XLIDE relationship (the seam)

XLIDE owns: the oracle (Excel/VBE ground-truth generation), the evidence
pipeline, and the data package below. Python owns: pure analysis only.

XLIDE exports a versioned **data package** that is the single contract surface.
Python pins a version of it and consumes it verbatim. Do not reimplement any of
this data in Python; load it.

| Asset | XLIDE source | Python consumes as |
|---|---|---|
| Oracle case corpus | `syntax_corpus/oracle/vbe_oracle_cases.json` (v1, 397 cases) | pytest parametrize fixtures; ground truth for present/absent diagnostics |
| Diagnostic provenance audit | `syntax_corpus/diagnostic_influence_audit.json` (117 diagnostics: 71 vbe-oracle-verified, 46 spec-derived) | single source of truth for which rules ship and their no-FP gating; the `notes` field becomes the rule docstring and gate |
| Rule metadata catalog | `src/analyzer/diagnostics/ruleMetadata.ts` (117 `code:` entries) | `rule_metadata.json` -> dataclasses (code, title, severity, category, diagnosticKind, specReference, confidence, suppressionScopes) |
| MS-VBAL verification map | `docs/spec/MS-VBAL.verification-map.md` | spec-reference annotations and a coverage report |
| Corpus provenance | `syntax_corpus/corpus_provenance.json` | CI integrity check |
| Managed backlog | `syntax_corpus/managed_backlog.md` | decision record: why a candidate is NOT ported |
| VBA runtime metadata | `src/analyzer/runtime/vbaRuntime.ts` | `vba_runtime.json` (host-agnostic; the only metadata pure-VBA needs) |
| Excel host tables | `src/analyzer/host/excelReferenceMembers.ts` + `reference/excel/json/*.json` (1029) + `reference/office/json/*.json` (511) | `excel_host.json` + `office_constants.json` (optional host layer) |

Oracle case shape: `{id, description, provenance, expected, mode, source,
evidencePhase, diagnosticMeaning}`. Audit entry shape: `{code, status,
sourceOfTruth, corpusInfluence, assertedOracleCases[], observeOnlyOracleCases[],
pendingLegacyCorpusFiles[], diagnosticKind, notes}`.

Proposed export artifact (XLIDE side, one-time setup): a `pyvba-data-package/`
directory bundling `vbe_oracle_cases.json`, `diagnostic_influence_audit.json`,
`rule_metadata.json`, `vba_runtime.json`, `excel_host.json`,
`office_constants.json`, and a top-level `manifest.json` carrying a semver, a
sha256 per file, the 117 code ids, the 85 rule ids, and an explicit
code-to-function map (see Risk 8). Python vendors this under
`pyvbaanalysis/data/` and pins the version. Until XLIDE builds the export, the
porting agent reads the raw files above directly and treats XLIDE's TypeScript as
the executable spec.

The no-false-positive rule is inherited and non-negotiable: a Python rule ships
only when it has a `diagnostic_influence_audit.json` entry with status in
{spec-derived, vbe-oracle-verified}. Never invent a rule. Always port a rule
together with its `assertedOracleCases` as fixtures. Copy the audit `notes`
verbatim as the rule's gate and docstring.

---

## 3. Module layout

Idiomatic Python that mirrors XLIDE's separation of concerns (UM-01, UM-02).
Use `Enum` for every `Kind`/`Severity` discriminant, `@dataclass(frozen=True)`
for nodes with `isinstance` dispatch (no TS unions), `frozenset`/`tuple` for
read-only data, `functools.lru_cache` or a dict where XLIDE uses a WeakMap,
`typing.Protocol` for visitor/tracker interfaces, and case-insensitive lookups
via lowercased keys (VBA is case-insensitive).

```
pyvbaanalysis/
  __init__.py                  # public API: analyze_module, parse_module, tokenize, ProjectIndex, analyze_workbook
  reader/
    workbook.py                # read VBA modules from Excel .xlsm/.xlsb/.xlam/.xls via pyOpenVBA -> source; the ONLY pyOpenVBA import
  lexer/
    token_kinds.py             # TokenKind, TriviaKind (Enum); VbaToken, Trivia (dataclass)
    tokenize.py                # tokenize(), tokenize_cached(), date-literal validator
    token_helpers.py           # match_paren_from, split_top_level_token_groups, statement_tokens, ...
    trivia.py                  # leading trivia, line-continuation
    keyword_table.py           # canonical_keyword, RESERVED_*; loads keyword data
    stripped_lines.py
  constants/
    integer_constant_expression.py   # evaluate, resolve_raw, cycle detection
  conditional/
    conditional_compilation.py # ConditionalActivity (Enum), stack-replay tracker, evaluator
  parser/
    nodes.py                   # all Node/Expr dataclasses; NodeKind/ExprKind Enums; Span
    parse_module.py            # parse_module() entry + LRU
    parse_expression.py        # ExpressionParser (precedence climbing)
    parser_state.py            # LogicalStatement, StatementCursor, split_logical_statements
    type_declaration_suffix.py
    fixed_length_string.py
  symbols/
    symbol_model.py            # VbaSymbol, VbaSymbolKind, signatures, formatters
    build_module_symbols.py
    name_resolution.py         # resolve_bare_identifier_binding (MS-VBAL 5.3)
    project_index.py           # ProjectIndex class
  types/
    type_inference.py          # signature tables, infer_argument_type, normalize_type
  diagnostics/
    model.py                   # VbaDiagnostic, DiagnosticSeverity/Category/EvidenceKind Enums
    analyze_module.py          # analyze_module(source, options) orchestrator
    context.py                 # RulePassContext (shared, memoized)
    registry.py                # ordered rule list (85 entries); output order is a contract
    rule_metadata.py           # loads rule_metadata.json (117 codes)
    call_extraction.py
    walker.py / exprwalk.py / dataflow.py   # shared walks + branch-merged lattice
    rules/                     # one module per family, mirrors src/analyzer/diagnostics/rules/**
      lexical.py duplicates.py declarations.py call_arity.py
      argument_types.py argument_shape.py runtime_values.py assignments.py
      arrays.py type_of_is.py binary_operand_scalar.py object_state.py
      undeclared.py module_kind.py expressions.py control_flow.py
      numeric_literals.py structural.py shared.py
  runtime/
    vba_runtime.py             # loads vba_runtime.json (host-agnostic)
  host/                        # OPTIONAL Excel layer; pure-VBA analysis runs without it
    host_model.py object_model.py
    data/excel_host.json office_constants.json
  data/                        # the vendored, version-pinned XLIDE export
    manifest.json vbe_oracle_cases.json diagnostic_influence_audit.json rule_metadata.json
tests/
  ... mirrors tests/diagnostics/** (18 family suites)
  conftest.py                  # loads the data package
  differential/                # TS-vs-Python golden harness (see Section 5)
```

Do not over-split (UM-04): one module per rule family, not one per rule. Do not
let any file become a monolith (UM-03): the 8900-line `excelReferenceMembers.ts`
must land as JSON data plus a thin resolver, never as a giant `.py`.

The reader layer (`reader/workbook.py`) is the only place pyOpenVBA is imported.
It turns an Excel workbook (.xlsm/.xlsb/.xlam/.xls) into module source text and
feeds the pure analyzer; `analyze_module(source)` and everything below it stays
stdlib-only (UM-01), so the core is testable from plain strings without any file. The
import name is most likely `pyopenvba` (per the wheel `pyopenvba-3.0.1`); confirm
at first use.

---

## 4. Parity inventory (what to port)

Tags: CORE = analysis-core, must port. HOST = Excel-specific, port as data.
EDITOR = out of scope. Size to port: S/M/L.

- Lexer (`src/analyzer/lexer/**`, `constants/integerConstantExpression.ts`):
  `tokenize` with full round-trip fidelity (CORE, S); 14 token kinds plus trivia
  (CORE, S); line-continuation trivia (CORE, S); statement-start context for Rem
  and directives (CORE, S); canonical keyword casing and the keyword tables
  RESERVED_*, CONTEXTUAL_KEYWORDS, OPERATOR_IDENTIFIERS, etc. (CORE data, S);
  integer/hex/octal/float literals with type suffixes `%&^!#@$` (CORE, S);
  date-literal ambiguous-grammar validator (CORE, M, subtle, see Risk 6);
  integer-constant-expression evaluator with cycle detection (CORE, M).
- Conditional compilation (`conditional/conditionalCompilation.ts`): directive
  index, `#Const` extraction, the default constants (VBA7=true, Win64=true,
  Win32=false, Mac=false), the Or/And/Not/comparison evaluator that returns a
  value or unknown, and the stack-replay activity tracker with O(log n) span
  queries (CORE, M). Unknown stays unknown; never guess a branch.
- Parser/AST (`parser/**`): `parse_module` entry (CORE, M); the full node set
  (24 NodeKind, 10 ExprKind) with every field (CORE, M); the precedence-climbing
  expression parser with bang/dot member access, New/AddressOf/TypeOf..Is, and
  named/omitted args (CORE, M); logical-statement splitting per MS-VBAL 3.3.1
  (CORE, S); IfBranch flow modeling (CORE, S); type-suffix mapping (CORE data,
  S); error tolerance (never throw; emit ParseDiagnostic and best-effort nodes;
  dangling-directive auto-close; stray-closer recovery) (CORE, M); the
  structured-vs-raw fallback (Assignment/Call/blocks structured; goto/label/
  return/exit raw) (CORE, M, see Risk 4).
- Symbols and scope (`symbols/**`): the symbol model (16 kinds, visibility,
  attributes) (CORE, S); `build_module_symbols` honoring conditional activity
  (CORE, M); signature model and formatters (CORE, S); `ProjectIndex` cross-
  module class (visibility, type/member surfaces, definition resolution,
  duplicate detection) (CORE, L); bare-identifier resolution per MS-VBAL 5.3 with
  local/module/project plus ambiguous/unresolved tri-state and shadowing (CORE,
  M); member-surface resolution with Property Get/Let/Set merge, `VB_UserMemId=0`
  default member, and `Implements` inheritance (CORE, M).
- Type inference (`diagnostics/typeInference.ts`): signature tables with
  memoization (CORE, L); `is_known_scalar_type`, `normalize_type` (CORE, S);
  `infer_argument_type` over literals/identifiers/call-returns/member-access/New
  with const and enum folding (CORE, L); `resolve_known_object_assignment_type`
  (CORE, M).
- Dataflow (`diagnostics/dataflow.ts`, `walker.ts`, `exprWalk.ts`): one shared
  statement pass and one shared expression pass per body (CORE, M); the
  straight-line lattice walk (CORE, M); `walk_branch_merged_body`, the meet-
  toward-unknown join over balanced If arms (CORE, L, hardest algorithm, see
  Risk 3); the object-var-not-set and unallocated-dynamic-array state trackers
  (CORE, M).
- Diagnostics engine (`diagnostics/analyzeModule.ts`, `analysisContext.ts`,
  `registry.ts`, `ruleMetadata.ts`, `callExtraction.ts`): `analyze_module(source,
  options)` orchestrator (CORE, M); the options input contract (project
  visibility, host model, conditional compilation, parsed module, severity
  overrides) (CORE, S); the shared/memoized `RulePassContext` (CORE, M); the
  85-entry ordered registry with run/procedureStatements/procedureExpressions
  forms and per-rule buffered flush, where output order is a contract (CORE, M);
  the 117-code metadata (CORE data, S); the severity model with off/downgrade
  overrides and the suppression-scope vocabulary (CORE, S); call extraction with
  arity and named-arg detection (CORE, M).
- Rules (`diagnostics/rules/**`, 17 files, 85 functions, 117 codes): lexical,
  duplicates, declarations, call arity, argument types, argument shape, runtime
  values, assignments, arrays, type-of-is, binary-operand-scalar, object state,
  undeclared (member-not-found), module kind, expressions, control flow, numeric
  literals, plus the structural block-balance pass. One function can emit several
  codes (for example arrays: 8 functions -> 12 codes; control flow: 9 -> 14), so
  track parity per code, not per function.
- Host and runtime metadata: `vbaRuntime.ts` is host-agnostic and required
  (CORE data, M). `excelReferenceMembers.ts` (229 types, 5900+ members, 36 of
  them with hard diagnostics) plus the 1029 Excel and 511 Office reference JSON
  files are the Excel host layer (HOST, L). Port as JSON plus thin resolvers;
  replicate `scripts/generate-excel-reference-metadata.mjs` in Python so the host
  JSON regenerates from `reference/**`.

---

## 5. Port order and validation

The validation backbone is a TS-vs-Python differential harness. Build it first;
it makes every later milestone objectively checkable.

- M0: Data ingestion plus the differential harness. Vendor the data package into
  `pyvbaanalysis/data/`. Add a small Node script in XLIDE that dumps, for each
  oracle `source` and each `tests/diagnostics/**` fixture, the TS `tokenize`,
  `parse_module`, `build_module_symbols`, and `analyze_module` output to JSON.
  Python compares against these goldens. Validate: harness runs, goldens
  captured.
- M1: Lexer plus integer-const-expr. Validate: round-trip (trivia plus rawText
  reconstructs the source exactly) on all 397 oracle sources; token parity vs the
  TS dump and `vbaLexer.test.ts`. Round-trippability is the acceptance gate.
- M2: Parser/AST plus parser_state. Validate: AST shape and span parity vs the TS
  `parse_module` dump over the corpus, plus `parseExpression.test.ts` and
  `vbaParser.test.ts`; a never-throws property test on malformed input.
- M3: Conditional compilation. Validate: `vbaConditionalCompilation.test.ts`
  (VBA7/Win64/Win32/Mac plus project `#Const`; unknown stays unknown).
- M4: Symbols, ProjectIndex, name resolution. Validate: `vbaSymbolGraph.test.ts`,
  `vbaProjectIndexService.test.ts` (visibility, ambiguity preserved, shadowing);
  inactive-branch filtering vs M3.
- M5: Type inference, call extraction, engine skeleton (registry, context,
  severity, buffered flush, zero rules). Validate: signature-table and
  `infer_argument_type` parity vs the TS dump; `analyze_module` returns empty in
  registry order.
- M6: First rules: structural, lexical, duplicates, declarations, control flow,
  expressions (token/structure-heavy, low type dependence; about 46 functions).
  Validate: per-family suites plus every `assertedOracleCase` for these codes
  fires or stays silent per the audit `status`.
- M7: Dataflow plus object_state plus arrays (needs `walk_branch_merged_body`).
  Validate: object-var-not-set and unallocated-array oracle cases including the
  balanced-If merge controls.
- M8: Type and call rules: argument types, argument shape, call arity,
  assignments, runtime values, type-of-is, binary-operand-scalar, numeric
  literals. Validate: the runtime-error 13/94/6 and arity (PCEC_008) oracle cases
  plus their compile-valid controls.
- M9: Project and host rules: undeclared (member-not-found), module kind, plus
  the VBA runtime metadata; Excel host as an optional extra. Validate: cross-
  module fixtures; member-not-found fires only on the 36 exhaustive host types.
- M10: Full parity sweep. Run all 397 oracle cases and all family suites through
  the differential harness. Acceptance: registry output order matches, every
  shipped code is reachable, zero diffs vs TS on control cases.

At each milestone, follow the agentic loop in the best-practices doc: understand,
plan the smallest coherent step, change with tests, validate with the harness and
pytest, report honestly (UM-08). "Done" means the harness and the relevant oracle
cases are green, not "looks right."

---

## 6. Sync plan (keeping Python in parity as XLIDE expands)

Shared contract. XLIDE is the sole owner of the oracle, the evidence pipeline,
and the data package. It exports a versioned `pyvba-data-package/` (manifest with
semver, per-file sha256, the 117 code ids, the 85 rule ids, and the
code-to-function map). Python pins a package version. Python never regenerates
oracle data and never invents a rule that lacks an audit entry.

Drift detection (Python CI). A `test_manifest_parity.py` that fails when:

1. A code in the audit with status in {spec-derived, vbe-oracle-verified} has no
   Python rule emitting it (unported code).
2. A Python rule emits a code absent from the manifest (orphan or divergent
   code).
3. A manifest `assertedOracleCase` id has no corresponding Python fixture
   assertion (untested case).
4. The vendored package sha256 does not match the manifest (stale vendor).

This turns the audit into an executable parity gate.

Porting a new XLIDE rule. (a) XLIDE adds the rule, its oracle cases, the audit
entry, and the metadata, then re-exports the data package with a version bump.
(b) Python bumps the pin; drift CI flags the new unported code red. (c) Port the
rule function into the matching `rules/*.py`, copying the audit `notes` verbatim
as the no-FP gate and docstring. (d) The rule's `assertedOracleCases` become its
fixtures automatically. (e) Drift CI goes green. Rules are always ported with
their oracle cases as fixtures, never code without evidence.

Avoiding divergence. Single source of truth for evidence: the data package.
Python holds no independent oracle or spec. Registry order is part of the
contract and is snapshot-tested. Keep the differential harness in CI on a
schedule so a TS behavior change surfaces as a Python diff even before a manifest
bump.

---

## 7. Conventions and operating rules

- Inherit `docs/agentic_ai_programming_best_practices.md` in full. Highlights
  that bite here: small focused changes (RG-01); tests as contracts (RG-04);
  simplicity over generality (RG-05); ground every claim against the XLIDE source
  (RG-18); honest status (UM-08); stop and ask when blocked (UM-09).
- No-false-positive discipline is the prime directive. If a construct is not
  provably wrong via the audit plus its oracle/spec evidence, the analyzer stays
  quiet. A false positive is a worse failure than a missed diagnostic.
- Dependencies (RG-10): pyOpenVBA is the only runtime dependency, and only the
  reader layer imports it. Everything else is the Python standard library. Do not
  add a runtime dependency; if you think you need one, stop and ask (UM-09).
- Plain ASCII only (UM-07). Use `->` not an arrow glyph, straight quotes, no em
  dashes, no emoji, spelled-out spec references.
- Mark generated and vendored content. `pyvbaanalysis/data/` and
  `host/data/` are vendored from XLIDE; do not hand-edit them. The host JSON is
  generated from `reference/**`; regenerate, do not edit.
- The XLIDE TypeScript is the executable spec during the port. When the harness
  shows a diff, XLIDE is right unless you can prove a TS bug, in which case fix it
  in XLIDE first and re-export, never fork the behavior in Python.
- Determinism. Same input plus same data-package version yields the same
  diagnostics in the same order, every run.

---

## 8. Risks and open decisions (flag to the user)

1. Excel host scope for v1. The host layer (excelReferenceMembers plus 1029+511
   reference JSON) is large and Excel-specific. Recommendation: ship the host-
   agnostic core first (lexer through M8 plus vbaRuntime), then add the Excel host
   layer (M9) to light up member-not-found and host type narrowing. Decide
   whether full Excel parity is required for the first usable release or can
   follow.
2. Node typing model. Use a frozen `@dataclass` hierarchy with `isinstance`
   dispatch (recommended) rather than a single tagged dataclass. This affects
   every walker and rule; decide once, up front.
3. `walk_branch_merged_body` is the hardest single algorithm and underpins the
   no-FP soundness of object-var-not-set and unallocated-array. Gate it behind its
   full oracle-control set before enabling M7 and M8.
4. Parser structured-vs-raw fallback. Rules branch on whether a statement is a
   structured Assignment/Call or a raw Statement. A Python parser that structures
   more or less than TS silently changes which rules fire. The differential AST
   harness must assert raw/structured parity exactly.
5. Performance. Target under 1 second for a 10k-line module. WeakMap memo becomes
   lru_cache or dict; Python is slower per token. If the lexer hot path is too
   slow on large modules, consider a native extension, but only after correctness
   parity.
6. Date-literal ambiguous grammar. The "try all candidate readings, accept if any
   matches exactly" routine is the trickiest lexer code. Port it verbatim and
   validate against the `vbaLexer.test.ts` date cases first.
7. Resolver seams that look like editor code but are CORE.
   `resolve_exact_member_completion` and `resolve_exhaustive_member_surface` live
   under `completion/` and `typeInference.ts`, but member-not-found,
   scalar-member-access, and argument-object-type rules depend on them. Port the
   resolver core; drop only the completion-array/UI wrapping. Misclassifying these
   as editor code would silently disable real rules.
8. Code-to-function map. 117 codes map onto 85 functions (one function can emit
   several codes). The manifest must carry an explicit code-to-function map so
   drift detection is per code. It is implicit in `registry.ts` plus
   `ruleMetadata.ts` today; make it explicit in the data package.

---

## 9. Key XLIDE anchors for the executing agent

- Registry and output order: `src/analyzer/diagnostics/registry.ts`
- Rule metadata (117 codes): `src/analyzer/diagnostics/ruleMetadata.ts`
- Rule families (17 files): `src/analyzer/diagnostics/rules/*.ts`
- Public entry and result type: `src/analyzer/diagnostics/analyzeModule.ts`
- Dataflow and walkers: `src/analyzer/diagnostics/{dataflow,walker,exprWalk}.ts`
- Evidence: `syntax_corpus/diagnostic_influence_audit.json`,
  `syntax_corpus/oracle/vbe_oracle_cases.json`
- Spec map: `docs/spec/MS-VBAL.verification-map.md`
- Host generation: `scripts/generate-excel-reference-metadata.mjs`,
  `scripts/generate-office-reference-metadata.mjs`
- XLIDE tests to mirror: `tests/diagnostics/**`, `tests/vba*.test.ts`

First action for the agent: read this file, then read the XLIDE registry,
ruleMetadata, and the audit JSON, then start M0 (vendor the data package and
stand up the differential harness). Do not write a rule before its oracle cases
are available as fixtures.
