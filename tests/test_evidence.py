"""M0: the vendored XLIDE evidence corpus loads and matches its manifest."""

from __future__ import annotations

import hashlib

from pyvbaanalysis.evidence import DATA_DIR, load_audit, load_manifest, load_oracle_cases


def test_oracle_cases_load() -> None:
    cases = load_oracle_cases()
    manifest = load_manifest()
    assert len(cases) == manifest["oracleCaseCount"] == 397
    assert all(c.modules for c in cases)
    assert all(m.source for c in cases for m in c.modules)
    assert all(c.expected in ("rejected", "accepted", "observe") for c in cases)
    assert len({c.id for c in cases}) == len(cases)  # unique ids


def test_audit_loads_and_matches_manifest() -> None:
    audit = load_audit()
    manifest = load_manifest()
    assert len(audit) == manifest["diagnosticCodeCount"] == 117
    assert {a.code for a in audit} == set(manifest["diagnosticCodes"])
    assert all(a.status in ("vbe-oracle-verified", "spec-derived") for a in audit)


def test_manifest_integrity() -> None:
    # The vendored files have not drifted from the manifest checksums.
    manifest = load_manifest()
    for name, meta in manifest["files"].items():
        data = (DATA_DIR / name).read_bytes()
        assert hashlib.sha256(data).hexdigest() == meta["sha256"], name
