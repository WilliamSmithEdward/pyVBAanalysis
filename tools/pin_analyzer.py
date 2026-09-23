"""Make a read-only copy of the XLIDE checkout at one commit, to vendor data from.

The extractors read the neighbouring xlide_vscode checkout, which is a WORKING TREE
somebody is usually working in. Both syncs so far were blocked by that: on 2026-09-04
it was 192 commits stale AND carried uncommitted edits, so it could not be
fast-forwarded without touching someone else's work, and the sync was done from an
improvised worktree that left git state behind in their repository.

So a sync vendors from a PIN instead: a copy of that repository at a chosen commit,
made once, never edited, and thrown away afterwards. Copying reads the source
repository and writes nothing to it, and it carries committed objects only, so work
in progress there is neither included nor disturbed.

Prints the root to vendor from. Point the extractors at it:

    python tools/pin_analyzer.py --ref v10.5.0
    XLIDE_ROOT=<printed path> npx -y tsx tools/extract_rule_metadata.mjs

or let tools/vendor_data.py do both.

XLIDE_ROOT is the checkout ROOT, not its ``src``: this repository vendors JSON out
of ``syntax_corpus/`` as well as TypeScript out of ``src/analyzer/``. xlide_vbide's
XLIDE_ANALYZER_ROOT points one level deeper because the engine only needs the source.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent
# Under artifacts, which is not in source control: a pin is a build input,
# reproducible from the ref at any time, and far too large for history.
_PIN_PARENT = _REPO_ROOT / "artifacts" / "analyzer-pin"
# Proves the pin is the analyzer repository and not some other checkout.
_MARKER = Path("src") / "analyzer" / "diagnostics" / "ruleMetadata.ts"


def _git(*args: str, cwd: Path | None = None) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=str(cwd) if cwd is not None else None,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise SystemExit(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout.strip()


def pin(ref: str, source: Path, force: bool = False) -> Path:
    """Copy ``source`` at ``ref`` into artifacts/ and return the pinned root."""
    if not (source / ".git").exists():
        raise SystemExit(f"No git repository at {source}, so there is nothing to pin.")

    # RESOLVED IN THE SOURCE, so the pin folder is named by commit whatever spelling
    # was asked for, and two names for one commit do not make two clones.
    commit = _git("rev-parse", "--verify", f"{ref}^{{commit}}", cwd=source)
    short = commit[:7]
    pin_root = _PIN_PARENT / short

    if (pin_root / _MARKER).is_file() and not force:
        print(f"Pin already at {pin_root} ({short})", file=sys.stderr)
        return pin_root

    if pin_root.exists():
        shutil.rmtree(pin_root, ignore_errors=True)
    pin_root.mkdir(parents=True, exist_ok=True)

    # THE OBJECT STORE IS COPIED, NOT CLONED. `git clone` reads the source's `.git` as
    # a path of its own and refuses it as "dubious ownership" whenever the two
    # checkouts were made by different accounts, and the remedy git suggests is an
    # entry in the GLOBAL config, a shared setting changed to do one build. A copy
    # needs no trust: the pin's `.git` is then this user's own. It carries committed
    # objects only, so work in progress in the source is neither included nor
    # disturbed, and the checkout below is pristine.
    shutil.copytree(source / ".git", pin_root / ".git")
    _git("config", "core.bare", "false", cwd=pin_root)
    # Vendored JSON must stay LF: a smudged checkout would hash CRLF into the
    # manifest and fail checksum verification on a Linux CI checkout.
    _git("config", "core.autocrlf", "false", cwd=pin_root)
    _git("checkout", "--quiet", "--force", commit, cwd=pin_root)

    dirty = _git("status", "--porcelain", cwd=pin_root)
    if dirty:
        raise SystemExit(
            f"The pin at {pin_root} is not pristine after checkout "
            f"({len(dirty.splitlines())} path(s) differ)."
        )
    if not (pin_root / _MARKER).is_file():
        raise SystemExit(f"The pin at {pin_root} has no {_MARKER}; is {source} the XLIDE repository?")

    described = _git("describe", "--tags", "--always", commit, cwd=pin_root) or short
    print(f"Pinned {described} ({short}) at {pin_root}", file=sys.stderr)
    return pin_root


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--ref", required=True, help="commit, tag or branch to pin (a sync pins a release tag)")
    parser.add_argument(
        "--source",
        type=Path,
        default=_REPO_ROOT.parent / "xlide_vscode",
        help="the XLIDE checkout to copy from (default: the sibling)",
    )
    parser.add_argument("--force", action="store_true", help="rebuild the pin even if it is already there")
    args = parser.parse_args()
    # The root goes to stdout alone, so a caller can capture it.
    print(pin(args.ref, args.source.resolve(), args.force))


if __name__ == "__main__":
    main()
