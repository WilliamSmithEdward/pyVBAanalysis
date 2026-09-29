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

Every push, pull request and release runs [security.yml](.github/workflows/security.yml):

- CodeQL with the `security-extended` queries, over the Python package and the
  GitHub Actions workflows.
- Semgrep with the `python`, `security-audit`, `secrets` and `github-actions`
  rule sets.
- pip-audit over the runtime dependencies.

Any finding fails the run, and a release is not published to PyPI until all
three pass. Each release carries the resulting `security-report.md` and the raw
SARIF and pip-audit output as assets. Dependabot keeps the dependencies and the
workflow actions current, and the run repeats weekly so new advisories surface
between releases.
