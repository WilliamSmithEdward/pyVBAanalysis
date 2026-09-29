# Security

## Supported versions

Security fixes go into the latest release on PyPI. Upgrade to it before
reporting a problem.

## Reporting a vulnerability

Report privately through GitHub:
[Security > Report a vulnerability](https://github.com/WilliamSmithEdward/pyVBAanalysis/security/advisories/new).
Please do not open a public issue for a vulnerability.

Include the version, what you ran, the input that triggers the problem (a
minimal module or file if you can share one), and what happened. You will get a
reply within a week.

## What counts

pyVBAanalysis reads Office files and VBA source and never runs them. A file
that makes the reader or the analyzer crash, hang, use unbounded memory, write
outside where it was asked to, or run code is a vulnerability. A wrong or
missing diagnostic is an ordinary bug; open an issue for it.

## How the code is checked

Every push, pull request and release runs [security.yml](.github/workflows/security.yml)
and [malware-scan.yml](.github/workflows/malware-scan.yml), and both run daily
as well:

- CodeQL with the `security-extended` queries, over the Python package and the
  GitHub Actions workflows.
- Semgrep with the `python`, `security-audit`, `secrets` and `github-actions`
  rule sets.
- pip-audit over the runtime dependencies.
- A malware scan of the tracked files and the built wheel and sdist: ClamAV,
  with signatures updated on every run and macro and heuristic alerts on, and
  YARA-X with the full rule set of [YARA Forge](https://github.com/YARAHQ/yara-forge),
  which collects the public YARA rule repositories.

Any finding fails the run, and a release is not published to PyPI until every
check passes. Known acceptable malware-scan findings are listed, each with its
reason, in [security/malware-allowlist.toml](security/malware-allowlist.toml);
there are none so far. Each release carries the resulting `security-report.md`
and `malware-report.md`, the latter naming the ClamAV signature version and the YARA Forge release it scanned with,
and the raw results as assets.

Known acceptable findings in the other checks are suppressed at the line, each
with its reason:

- Semgrep `dynamic-urllib-use-detected` in `security/yara_forge_update.py`: the
  updater fetches release URLs the GitHub API returns, and refuses any URL that
  is not https on the GitHub API or the YARA Forge release download path.

Everything the checks run on is pinned: actions by commit SHA, the ClamAV and
Semgrep engines by image digest, pip-audit by hash, the YARA-X engine by version
and SHA-256, and the YARA Forge rules by release and SHA-256. Dependabot proposes
new versions of the first three once they are a week old. A weekly workflow
proposes the newest YARA Forge release a week old in a pull request, and the
Malware scan workflow scans that pull request before it can be merged.
