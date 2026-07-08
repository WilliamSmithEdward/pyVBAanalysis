"""Regenerate pyvbaanalysis/data/manifest.json from the vendored evidence files.

The manifest pins the data package: per-file sha256 and size, the oracle case
count, the audited diagnostic-code list, and the rule catalogue names. Re-run
after re-vendoring any of the three evidence files:

    python tools/generate_manifest.py <xlideVersion>

The xlideVersion argument is the upstream xlide_vscode release the files were
vendored from (e.g. 2.5.11).
"""

from __future__ import annotations

import hashlib
import json
import sys
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
    if len(sys.argv) != 2:
        sys.exit("usage: python tools/generate_manifest.py <xlideVersion>")
    xlide_version = sys.argv[1]

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
