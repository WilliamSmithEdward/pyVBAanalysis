# Notes for agents

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** publishing a GitHub release runs Publish, which uploads to PyPI and attaches the security and malware reports to the release.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->

## What this repository is

pyVBAanalysis is a Python port of the analyzer core of
WilliamSmithEdward/xlide_vscode ("upstream", "XLIDE"). The sync point is
whatever `pyvbaanalysis/data/manifest.json` says (`xlideCommit`), not what
the changelog claims; check a specific upstream change is present before
trusting a version label.

- The upstream checkout is `F:\GitHub\xlide\xlide_vscode` (the tools also fall
  back to the old sibling `..\xlide_vscode`). It is the owner's working tree:
  read it, never install into it, never modify it, never vendor from it. The
  owner edits it while you work, so a probe of the live tree can see a
  half-finished state. For anything you will cite, probe a clean export of a
  commit instead (see "Hunting analyzer faults").
- Vendored upstream data comes from a pinned checkout under
  `artifacts/analyzer-pin/<short-sha>`, made by `tools/vendor_data.py --ref
  vX.Y.Z`. It vendors evidence JSON with `git show`, never by copying a working
  tree (an autocrlf checkout turns it to CRLF and the checksums fail on CI).
  Vendored files must be plain ASCII with LF line endings.
- A third repository feeds this family: `F:\GitHub\pyOpenVBA` (the Office file
  format library). The ground-truth Office oracle is `F:\GitHub\pyVBAharness`.
- Re-fetch before claiming an upstream commit or change does not exist. Absence
  is the one claim that expires in minutes; say when you checked.

## Syncing and verifying a sync

CONTRIBUTING.md ("Checking a sync against upstream") documents
`tools/differential/harness.py` (record, replay, corpus, cases, projects,
unpatch). The decisive check for a sync is that behavioral differential, not
the test suite: the port's tests encode old behavior and stay green through a
widened upstream rule.

- Run record, replay, corpus and projects, over Office files of every host
  (Excel, Word, PowerPoint, Access) plus export folders, and `--host vb6` on
  VB6 folders. Anchor a zero before trusting it: undo a fix in memory and
  confirm the comparison reports it.
- Where the port reports more, it is a port bug. Where upstream reports more,
  it may be an upstream false positive; check before porting it.
- Measure performance as CPU time in fresh processes, alternating baseline and
  candidate, with the baseline from `git archive HEAD pyvbaanalysis` and each
  child run outside the repo printing `pyvbaanalysis.__file__`.
- Never pass a raw string (an If condition, a default, a Const value) to the
  cached statement tokenizer; lex it with `diagnostics/walker.py
  raw_expression_tokens`. `tests/test_analyzer_lexes_module_once.py` guards it.
- Body walks run on explicit stacks (Python stops at 1000 frames; upstream
  recurses). Convert any newly ported recursive walker;
  `tests/test_deep_nesting.py` fails on one that recurses per block.
- Shared caches and host models are read-only. Memoize per model with
  `identity_cache.IdentityLru`, never a bare `dict[id(model)]`.

## Port bugs and parity bugs

- A bug only the port has (upstream is right): you have standing approval to
  fix it. Write a failing test first, then run ruff, mypy (from `.venv`) and
  the full suite. `main` is PR-only: push a `fix/` branch rebased on
  `origin/main`, open a PR, and run
  `gh pr merge N --auto --squash --subject "..." --body-file FILE`.
- A bug the port shares with upstream (parity): file it upstream only and wait
  for the sync. Do not patch the port ahead of upstream, and do not offer to.
- Honoring XLIDE's `@xlide-analysis-disable-*` comments is deliberately parked.
  The port reads only `'@pyvba-ignore`. Do not implement or re-ask.
- The CLI default stays `--fail-level information` (owner's decision).
- Before a release, check `git status`: other sessions sometimes leave
  uncommitted work here, and the owner may want it in.

## Hunting analyzer faults

The owner's priority is runtime "gold": code the VBE compiles that raises
every time it runs, unreported. False positives are equally important.
Compile errors the analyzer misses are missing red squiggles and are filed
too. Hunting targets upstream's current `main`.

### The oracle

pyVBAharness (`pip pyvbaharness`) drives real Office: ExcelSession,
WordSession, PowerPointSession, AccessSession; `new_document`, `add_module`,
`compile_project` (returns the VBE dialog text), `run_macro` (VBA errors come
back as number and description). Measurements so far are Excel 16.0 64-bit,
build 20430.

- A case is `{label, host?, modules: [{name, type, source}], run:
  "Module1.Main"}`. `host` picks the session (excel by default, word,
  powerpoint, access). Only standard and class modules can be injected.
- `compile_project` reports only the first error: one hypothesis per case.
- Access writes `Option Compare Database` into new modules; strip it from
  injected source. PowerPoint runs on screen and will not start while
  PowerPoint is open. Other sessions may hold the per-app lock: wait and
  retry, never pass `exclusive=False`.
- pyVBAharness serializes 0.5 as ".5" and reports it as a runner error; the
  hunt's `leading_dot.py` recovers those verdicts.
- Cases that touch the outside world must clean up: close files (`Close` at
  the top), delete temp files and registry settings they create, and never
  reference a missing sheet or file in a formula (Excel opens a dialog). Do not
  use SendKeys. A crash of Excel (RPC failure) turns every later case into a
  harness error: drop those verdicts and re-measure (`drop_harness.py`).

### The hunt toolkit

The scripts are in `tools/hunt/`, each with a usage docstring at its top. They
read and write their data in the current folder: work in `artifacts/hunt/`
(gitignored, like the rest of `artifacts/`), and run them by path, for
example `python ../../tools/hunt/table3.py ORACLE FOUND`. Never commit corpus
data, oracle files or exported trees.

As of 2026-10-03 the measured corpus, the exported upstream trees and the
474 one-off topic generators (`gen_<topic>.py`, one per batch) are still in a
session scratchpad, not in the repository:
`C:\Users\William\AppData\Local\Temp\claude\F--GitHub-pyVBAanalysis\a5820d5e-30fe-411c-aedf-a3f9c40f3d5a\scratchpad\hunt\`.
Copy `all.json`, `all_oracle.json` and the batch files from there into
`artifacts/hunt/` to continue from it; if it is gone, the corpus has to be
measured again.

- `export_remote.py SHA DIR`: download a commit of xlide_vscode as a tarball
  through `gh api` into DIR (no fetch into the owner's checkout). Probes then
  run with `XLIDE_ROOT=DIR`. Trees are kept as `xlide_<sha>`.
- `wrapper_probe.mjs CASES.json`: run upstream's full wrapper over cases
  (`XLIDE_ROOT=<tree> npx -y tsx wrapper_probe.mjs ...`), one JSON line per
  module with its codes and messages. Honors `host`.
- `oracle_run.py CASES.json OUT.json`: run each case through pyVBAharness.
  It resumes by label, so delete the output to re-measure.
- `queue_oracle.py WAIT.json N BATCH...`: wait until an oracle file holds N
  verdicts, then run more batches (one Excel at a time).
- `table3.py ORACLE FOUND [--diffs]`, `fp_rules.py ORACLE FOUND`,
  `show_cases.py CASES ORACLE FOUND LABEL...`, `short_fn.py` and
  `short_lost.py`: line Excel's verdict up with findings, list false
  positives by rule, print sources, show the shortest misses or losses.
- A new topic batch is a small generator `gen_X.py` writing `X.json`, built
  from `std(label, decl, body, extra)` in `tools/hunt/gen_dtc.py`; keep it in
  `artifacts/hunt/` with its data. Module-level declarations go in `decl`,
  never after the procedures.
- Other tools: `export_remote.py`, `drop_harness.py` (drop crashed verdicts
  to re-measure), `bis_first.py` (first bad commit from per-commit runs),
  `show_span.py` (source around a finding in a corpus part),
  `late_comments.py` (comments posted on issues after they closed).

### The corpus

All measured batches are merged into one corpus (67,712 cases as of
2026-10-03): `all.json` plus `all_oracle.json`, labels `batch::label`.

- After adding batches: `python merge_batches.py all.json all_oracle.json`,
  then `python leading_dot.py all_oracle.json all_oracle_ld.json`, copy that
  back over `all_oracle.json`, then `python split_cases.py all.json 6`.
- Before naming a new batch X, run
  `python recover_batch.py all.json all_oracle.json --list X` and `ls X.json`
  and read the output before generating anything. Names like `oct`, `rnd` and
  `wpn` were already taken. If a batch file is overwritten anyway,
  `recover_batch.py all.json all_oracle.json X NEW` rebuilds it from the
  corpus. Also `ls` a generator's file name before writing a new one.

### Sweeps, bisects and pull requests

- Sweep a tree: `sh ../../tools/hunt/sweep6.sh xlide_<sha> TAG` (from
  `artifacts/hunt/`) runs the six corpus parts in parallel into
  `all_TAG.jsonl`. Compare two trees with
  `regress.py all_oracle.json OLD.jsonl NEW.jsonl` (regressions) and
  `gains.py` (gained and lost, by batch).
- After every upstream move: sweep the new head against the previous one,
  bisect each loss by exporting the commits in between, read the losses (a
  "loss" of a finding that was right for the wrong reason is a fix, not a
  regression), and file what is real.
- The compare API lists commits oldest first; the head is the last one, or ask
  `gh api repos/.../commits/main`.
- For an open PR: find its merge base
  (`gh api repos/.../compare/main...SHA --jq .merge_base_commit.sha`), sweep
  both, and comment the result on the PR. Re-sweep the baseline after merging
  new batches, or the comparison shows fake regressions. For refactor and
  performance PRs, diff the full findings per module (messages included),
  not only verdicts; they should be identical.
- Real code: `corpus_two_trees.py OLD NEW OUT.jsonl <parts>` over
  `realc/real.*.json`, `artifacts/differential/cases/office_part*.json`,
  `workbooks.json` and `vb6.json` diffs findings between two trees (about an
  hour). `real_all.py` runs one tree and `real_errs_tally.py` with
  `error_codes.py` tallies error-severity findings by rule for triage. Run one
  heavy background job at a time; two at once crawl.
- Real-code noise to expect: VB6 `.frm` files probed as raw text, test modules
  probed without the library modules they call, and stale module types in the
  2026-09-22 Office case files (built before the reader learned the VBE class
  base GUID).

### Fuzzers

Random case generators, each `python gen_X.py OUT.json N SEED`, labels
`<prefix><seed>/NNNN`: `gen_flowfuzz.py` (flz), `gen_flowfuzz2.py` (flx,
arrays and Collections), `gen_flowfuzz3.py` (fle, On Error in branches),
`gen_exprfuzz.py` and `gen_exprfuzz2.py` (xf, typed expressions into typed
targets), `gen_callfuzz.py` (cz, random signatures and arguments),
`gen_cfz.py` (kcz, callees that change their caller's state),
`gen_datefuzz.py` (dtz), `gen_strfuzz.py` (sfz), `gen_rgfuzz.py` (rgz),
`gen_shfuzz.py` (shz), `gen_declfuzz.py` (dz, valid declarations),
`gen_colfuzz.py` (clz, name collisions), `gen_gdfuzz.py` (gfz, random guards
over known locals).

- Every generated loop needs its own counter variable, and every GoTo must jump
  forward to a unique label. Nested loops sharing a counter can loop forever
  and hang Excel; check generated cases before running the oracle.
- Use a fresh seed after each upstream change to the rules a fuzzer covers.

### What counts

- Only what a static rule can prove. Errors that depend on workbook contents,
  machine state or the registry (a Find that finds nothing, a missing sheet,
  an unregistered ProgID, a missing program) are not misses.
- By design, upstream reports a line that raises even when it never surfaces:
  divisions under `On Error GoTo`, in helpers nothing calls, in
  Class_Terminate, in event handlers. Deliberate `Err.Raise` is not a miss.
- Print the full source before doubting a verdict; a filtered view hides lines.

## Filing issues and comments

You have standing permission to file issues and comment on any of the owner's
repositories (xlide_vscode, pyOpenVBA, pyVBAharness, pyVBAanalysis, ...).
Fixing code outside pyVBAanalysis is not covered.

- Post only through the hunt's `post_issue.py`
  (`python post_issue.py issue "TITLE" BODY.md` or
  `python post_issue.py comment N BODY.md`; `ISSUE_REPO=owner/repo` for
  another repository). It refuses lines over 78 characters and code spans
  split across a wrap.
- Write bodies with the Write tool, check `awk '{ if (length($0) > 78) print NR }'`,
  and edit with `gh issue edit --body-file` or
  `gh issue comment N --edit-last --body-file`.
- Plain ASCII, no em dashes, no AI credit. Claim only what was measured: the
  Excel build, the commit probed, exact values. Check every number in a draft
  against the oracle file before posting.
- Search open and closed issues before filing (`gh issue list --search`).
- A regression issue names the commit range, the counts (gained, lost), and
  pins each loss to its commit. Bisect before filing.
- A comment on a closed issue can go unread. Post follow-ups to closed issues
  as a new issue (or add them to an open tracking issue), not as a comment.

## Shell and tooling rules

- Never use a shell heredoc, and never run `python -` reading stdin. Write
  files with the Write tool, change them with the Edit tool, and run scripts
  from files. Never sed-edit Python string escapes.
- Never `taskkill` python.exe or node.exe broadly; other sessions run them.
  Stop your own background jobs by task id.
- Long jobs (oracle runs, real-code diffs) go to the background with a
  generous timeout; check progress from their logs.
- Pin a new dependency version only once it is a week old.
