# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/pyVBAanalysis/security/advisories/new).
Only the maintainer sees it. Include the `pyvbaanalysis` version, what you
ran, and the smallest input that shows it (a minimal module or Office file
if you can share one), with credentials and private data removed.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory
is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older releases are
not maintained separately; update when a fix ships.

## Scope

pyVBAanalysis reads Office files and VBA source and never runs them. A file
that makes the reader or the analyzer crash, hang, use unbounded memory,
write outside where it was asked to, or run code is a vulnerability. A wrong
or missing diagnostic is an ordinary bug; open an issue for it.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily, and again from the
Publish workflow for every release.

- **Code:** CodeQL with the `security-extended` queries, for Python and the
  GitHub Actions workflows, and Semgrep with the `p/python`,
  `p/security-audit`, `p/secrets` and `p/github-actions` rule sets. Results
  go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** pip-audit over the runtime dependency tree, from the
  hash-locked `.github/requirements/runtime.txt` that matches what a user's
  install resolves today. Any known vulnerability fails Security.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every tracked file and the wheel and sdist built from
  them. ClamAV runs with macro and heuristic alerts on and reports every
  match on a file, not only the first. YARA-X runs the full YARA Forge pack.
- **Fuzzing:** Atheris fuzzes the lexer, the parser and every diagnostic rule
  on arbitrary VBA source (`fuzz/fuzz_analyzer.py`), starting from the seeds
  in `tests/fuzz_corpus/vba`: the lexer must never raise and its tokens must
  round-trip to the source, and the parser and rules must never raise. The
  Fuzz workflow runs on every change to the package, the fuzz target or its
  corpus, for a minute per target, and daily for five. It is not a gate: a
  finding becomes a regression test with its fix.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks assume more than one maintainer, such as a second
  person approving every change, so a single-maintainer project cannot
  score full marks on them.

## Accepted findings

A malware finding is fixed, or accepted with a written reason in
[security/malware-allowlist.toml](security/malware-allowlist.toml). An entry
matches the scanner, the rule and a glob over the path, and an entry for a
single file also matches its SHA-256, so a changed file needs another
review. An entry that no longer matches is noted in the scan log and the
raw results; it does not yet fail the report. CodeQL and Semgrep have no
accepted list: any CodeQL finding fails Security, and a Semgrep finding can
be accepted only by a `nosemgrep` comment at the line, listed here with its
reason. zizmor keeps its exceptions in `.github/zizmor.yml` or inline
beside the line they excuse, each with its reason. The current entries:

- ClamAV `Heuristics.OLE2.ContainsMacros.VBA` on
  `tests/fixtures/PowerPointFixture.ppt`, pinned to its SHA-256: a legacy
  PowerPoint test file that holds VBA on purpose, since it is the only way
  to test the `.ppt` reader. It is left out of the published packages.
- zizmor `self-repository`, turned off in
  [.github/zizmor.yml](.github/zizmor.yml) until GitHub's documentation
  confirms the `$/` self-repository syntax for reusable workflows called
  from `publish.yml`.

No `nosemgrep` comment is in use.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the development tools to exact versions in
`pyproject.toml`, and the YARA-X engine and YARA Forge rules to a release
and its SHA-256. ClamAV's signatures change too often to pin, so freshclam
fetches and verifies them on every run. The Semgrep rule sets are fetched
from the Semgrep registry on every run.

Dependabot proposes updates to the GitHub Actions, the scanner images,
`pyproject.toml` and the lock files in `.github/requirements` once a
version is a week old, and at once for a security advisory. The Update YARA
rules workflow proposes new YARA pins each week. A minor or patch update,
and the YARA pull request, merges itself once CI, Security and Malware scan
pass; a third-party major version waits for review.

## Releases

Publishing a GitHub release starts the Publish workflow. It runs the tests,
ruff and mypy, builds the wheel and sdist, runs Security and Malware scan
on the release commit, and uploads to PyPI through trusted publishing only
when the build and both scans pass. Started by hand, it is a dry run that
publishes nothing.

The reports are attached only when both scans pass. Every release after
v2.3.1 carries `security-report.md` and `malware-report.md`, the latter
naming the ClamAV signature version and the YARA Forge release it scanned
with, the raw results (`security-sarif.tar.gz`, `malware-results.tar.gz`),
and the signed provenance bundle `pyvbaanalysis-<version>.sigstore.json`.
v2.3.1 carries the security report and its SARIF; earlier releases carry
no reports.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file, from PyPI or
from the GitHub release:

```
pip download pyvbaanalysis --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the GitHub release as
`pyvbaanalysis-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle pyvbaanalysis-<version>.sigstore.json`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
