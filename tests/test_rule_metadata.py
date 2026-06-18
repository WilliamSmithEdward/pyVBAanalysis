"""M5a: the diagnostic model + vendored rule catalogue (ruleMetadata.ts data parity)."""

from __future__ import annotations

import json

from pyvbaanalysis.diagnostics import (
    DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES,
    DIAGNOSTIC_RULES,
    DiagnosticCategory,
    DiagnosticEvidenceKind,
    DiagnosticSeverity,
    DiagnosticSuppressionScope,
    load_rule_metadata,
    rule_metadata_by_code,
)
from pyvbaanalysis.diagnostics.model import VbaDiagnostic
from pyvbaanalysis.evidence import DATA_DIR, load_audit, load_manifest
from pyvbaanalysis.parser.nodes import Span

# The two structural block-balance codes are emitted by the parser pass, not the
# rule catalogue, so they appear in the audit but not in rule_metadata.
_STRUCTURAL_ONLY_CODES = {"missing-block-closer", "unmatched-block-closer"}


def test_catalogue_loads_115_rules() -> None:
    rules = load_rule_metadata()
    assert len(rules) == 115
    assert rules == DIAGNOSTIC_RULES  # the eager catalogue matches a fresh load


def test_codes_are_unique() -> None:
    codes = [meta.code for meta in DIAGNOSTIC_RULES.values()]
    assert len(codes) == len(set(codes)) == 115


def test_fields_are_typed_enums() -> None:
    for meta in DIAGNOSTIC_RULES.values():
        assert isinstance(meta.default_severity, DiagnosticSeverity)
        assert isinstance(meta.category, DiagnosticCategory)
        assert isinstance(meta.diagnostic_kind, DiagnosticEvidenceKind)
        assert meta.source == "XLIDE"
        assert meta.confidence in ("high", "medium", "low")
        if meta.suppression_scopes is not None:
            assert all(isinstance(s, DiagnosticSuppressionScope) for s in meta.suppression_scopes)


def test_rule_metadata_by_code() -> None:
    by_code = rule_metadata_by_code()
    assert len(by_code) == 115
    assert by_code["unterminated-string"].rule_name == "unterminatedString"
    assert by_code["unterminated-string"].default_severity is DiagnosticSeverity.ERROR


def test_codes_align_with_audit() -> None:
    rule_codes = {meta.code for meta in DIAGNOSTIC_RULES.values()}
    audit_codes = {a.code for a in load_audit()}
    assert rule_codes <= audit_codes
    assert audit_codes - rule_codes == _STRUCTURAL_ONLY_CODES


def test_manifest_records_rule_metadata() -> None:
    manifest = load_manifest()
    assert manifest["ruleCount"] == 115
    assert "rule_metadata.json" in manifest["files"]
    assert sorted(manifest["ruleNames"]) == sorted(DIAGNOSTIC_RULES.keys())


def test_vendored_json_is_ascii_and_lf() -> None:
    raw = (DATA_DIR / "rule_metadata.json").read_bytes()
    assert raw.isascii()
    assert b"\r\n" not in raw  # LF only (clone-safe checksums)
    json.loads(raw)  # well-formed


def test_default_suppression_scopes() -> None:
    assert DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES == (
        DiagnosticSuppressionScope.BLOCK,
        DiagnosticSuppressionScope.MEMBER,
        DiagnosticSuppressionScope.MODULE,
    )


def test_sample_rule_metadata() -> None:
    meta = DIAGNOSTIC_RULES["duplicateProcedure"]
    assert meta.code == "duplicate-procedure"
    assert meta.default_severity is DiagnosticSeverity.ERROR
    assert meta.category is DiagnosticCategory.DECLARATION
    assert meta.diagnostic_kind is DiagnosticEvidenceKind.COMPILE_ERROR
    assert meta.vbe_compile_equivalent is True
    assert meta.spec_reference == "MS-VBAL 5.3"


def test_vba_diagnostic_shape() -> None:
    d = VbaDiagnostic(code="x", message="m", severity=DiagnosticSeverity.WARNING, span=Span(0, 1))
    assert d.spec_reference is None and d.data is None
    assert d.severity is DiagnosticSeverity.WARNING
