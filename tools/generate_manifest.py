"""Regenerate pyvbaanalysis/data/manifest.json from the vendored evidence files.

The manifest pins the data package: per-file sha256 and size, the oracle case
count, the audited diagnostic-code list, and the rule catalogue names. Re-run
after re-vendoring any of the three evidence files:

    python tools/generate_manifest.py <xlideVersion> [--commit <sha>]

The xlideVersion argument is the upstream xlide_vscode release the files were
vendored from (e.g. 10.5.0), and --commit the exact commit. tools/vendor_data.py
runs this as its last step, with the commit it pinned.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_DATA = _ROOT / "pyvbaanalysis" / "data"
_OUT = _DATA / "manifest.json"

_FILES = (
    "vbe_oracle_cases.json",
    "diagnostic_influence_audit.json",
    "rule_metadata.json",
)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("xlide_version", help="the upstream release the files were vendored from")
    parser.add_argument(
        "--commit",
        help=(
            "the XLIDE commit vendored from. A version string is a label somebody typed; "
            "the commit is what makes the pin reproducible, so record it when it is known."
        ),
    )
    args = parser.parse_args()
    xlide_version = args.xlide_version

    files: dict[str, dict[str, object]] = {}
    for name in _FILES:
        raw = (_DATA / name).read_bytes()
        files[name] = {"sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)}

    cases = json.loads((_DATA / "vbe_oracle_cases.json").read_text(encoding="utf-8"))
    audit = json.loads((_DATA / "diagnostic_influence_audit.json").read_text(encoding="utf-8"))
    rules = json.loads((_DATA / "rule_metadata.json").read_text(encoding="utf-8"))

    manifest = {
        "sourceRepo": "WilliamSmithEdward/xlide_vscode",
        "xlideVersion": xlide_version,
        **({"xlideCommit": args.commit} if args.commit else {}),
        "files": files,
        "oracleCaseCount": len(cases["cases"]),
        "diagnosticCodeCount": len(audit["diagnostics"]),
        "diagnosticCodes": sorted(d["code"] for d in audit["diagnostics"]),
        "ruleCount": len(rules),
        "ruleNames": sorted(rules.keys()),
    }
    _OUT.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8", newline="\n")
    print(
        f"wrote {_OUT.relative_to(_ROOT)} - xlide {xlide_version}, "
        f"{manifest['oracleCaseCount']} cases, {manifest['diagnosticCodeCount']} codes, "
        f"{manifest['ruleCount']} rules"
    )


if __name__ == "__main__":
    main()
