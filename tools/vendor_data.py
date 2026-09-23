"""Vendor every data file from one pinned XLIDE commit, then write the manifest.

One command so a sync cannot half-happen: the generated models, the runtime tables,
the event catalogue and the evidence corpus all come from the same commit, and the
manifest records that commit rather than a version string somebody typed.

    python tools/vendor_data.py --ref v10.5.0

The pin is made first (tools/pin_analyzer.py), so nothing here reads the sibling
working tree. Pass --root to vendor from a checkout you have already prepared.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

from pin_analyzer import _git, pin

_REPO_ROOT = Path(__file__).resolve().parent.parent
_DATA = _REPO_ROOT / "pyvbaanalysis" / "data"

# Extractors, in the order a reader would want them reported.
_EXTRACTORS = (
    "extract_rule_metadata.mjs",
    "extract_runtime_tables.mjs",
    "extract_event_definitions.mjs",
    "extract_host_model.mjs",
)

# Evidence files, copied out of git rather than off disk. A checkout can smudge
# them to CRLF, the manifest then hashes CRLF bytes, and CI's Linux checkout fails
# checksum verification. `git show` hands back the committed bytes either way.
_EVIDENCE = {
    "vbe_oracle_cases.json": "syntax_corpus/oracle/vbe_oracle_cases.json",
    "diagnostic_influence_audit.json": "syntax_corpus/diagnostic_influence_audit.json",
}


def _run_extractors(root: Path) -> None:
    env = dict(os.environ, XLIDE_ROOT=str(root))
    for name in _EXTRACTORS:
        print(f"==> {name}", file=sys.stderr)
        result = subprocess.run(
            ["npx", "-y", "tsx", str(Path("tools") / name)],
            cwd=str(_REPO_ROOT),
            env=env,
            shell=True,
        )
        if result.returncode != 0:
            raise SystemExit(f"{name} failed")


def _vendor_evidence(root: Path, commit: str) -> None:
    for target, path in _EVIDENCE.items():
        blob = subprocess.run(
            ["git", "-C", str(root), "show", f"{commit}:{path}"],
            capture_output=True,
        )
        if blob.returncode != 0:
            raise SystemExit(f"could not read {path} at {commit}: {blob.stderr.decode(errors='replace').strip()}")
        raw = blob.stdout
        if b"\r\n" in raw:
            raise SystemExit(f"{path} came back with CRLF; the manifest would not verify on Linux")
        if not raw.isascii():
            raise SystemExit(f"{path} is not plain ASCII; the evidence files are ASCII by contract")
        (_DATA / target).write_bytes(raw)
        print(f"==> vendored {target} ({len(raw)} bytes)", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--ref", help="XLIDE commit, tag or branch to vendor from")
    parser.add_argument("--root", type=Path, help="an XLIDE checkout to use as-is instead of pinning")
    parser.add_argument("--source", type=Path, default=_REPO_ROOT.parent / "xlide_vscode")
    args = parser.parse_args()

    if args.root is not None:
        root = args.root.resolve()
    elif args.ref:
        root = pin(args.ref, args.source.resolve())
    else:
        raise SystemExit("pass --ref to pin a commit, or --root to use a prepared checkout")

    commit = _git("rev-parse", "HEAD", cwd=root)
    # The manifest has always carried the bare version, "6.2.0" rather than the
    # tag's "v6.2.0".
    described = _git("describe", "--tags", "--always", commit, cwd=root) or commit[:7]
    described = described.removeprefix("v")

    _run_extractors(root)
    _vendor_evidence(root, commit)

    print(f"==> manifest for {described} ({commit[:7]})", file=sys.stderr)
    subprocess.run(
        [sys.executable, str(Path("tools") / "generate_manifest.py"), described, "--commit", commit],
        cwd=str(_REPO_ROOT),
        check=True,
    )


if __name__ == "__main__":
    main()
